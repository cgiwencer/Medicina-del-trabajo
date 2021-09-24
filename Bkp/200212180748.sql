/*
MySQL Backup
Source Server Version: 5.1.30
Source Database: medtrabajo
Date: 12/02/2020 18:07:48
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
--  Table structure for `empresas`
-- ----------------------------
DROP TABLE IF EXISTS `empresas`;
CREATE TABLE `empresas` (
  `empid` int(6) NOT NULL AUTO_INCREMENT,
  `empdes` varchar(250) DEFAULT NULL,
  PRIMARY KEY (`empid`)
) ENGINE=InnoDB AUTO_INCREMENT=136 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `impbol`
-- ----------------------------
DROP TABLE IF EXISTS `impbol`;
CREATE TABLE `impbol` (
  `NueId` int(6) DEFAULT NULL,
  `NueNom` varchar(250) DEFAULT NULL,
  `NueSex` varchar(10) DEFAULT NULL,
  `NueFeN` date DEFAULT NULL,
  `NueEda` int(3) DEFAULT NULL,
  `NueTeI` varchar(20) DEFAULT NULL,
  `NueTer` varchar(20) DEFAULT NULL,
  `EmpDes` varchar(250) DEFAULT NULL,
  `NueFeI` date DEFAULT NULL,
  `NueFeS` date DEFAULT NULL,
  `NueFeP` date DEFAULT NULL,
  `NueFeM` date DEFAULT NULL,
  `NueNR` varchar(10) DEFAULT NULL,
  `NueNR1` varchar(10) DEFAULT NULL,
  `NueNR2` varchar(10) DEFAULT NULL,
  `NueNR3` varchar(10) DEFAULT NULL,
  `NueNR4` varchar(10) DEFAULT NULL,
  `NueEst` int(1) DEFAULT NULL,
  `UsuRes1` int(4) DEFAULT NULL,
  `ProEsRx` int(1) DEFAULT NULL,
  `ProEsLa` int(1) DEFAULT NULL,
  `ProEsMe` int(1) DEFAULT NULL,
  `UsuRes2` int(4) DEFAULT NULL,
  `NueObs` mediumtext
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `nuevos`
-- ----------------------------
DROP TABLE IF EXISTS `nuevos`;
CREATE TABLE `nuevos` (
  `NueId` int(6) NOT NULL AUTO_INCREMENT,
  `NueNom` varchar(250) DEFAULT NULL,
  `NueSex` varchar(10) DEFAULT NULL,
  `NueFeN` date DEFAULT NULL,
  `NueEda` int(3) DEFAULT NULL,
  `NueTeI` varchar(20) DEFAULT NULL,
  `NueTer` varchar(20) DEFAULT NULL,
  `EmpDes` varchar(250) DEFAULT NULL,
  `NueFeI` date DEFAULT NULL,
  `NueFeS` date DEFAULT NULL,
  `NueFeP` date DEFAULT NULL,
  `NueFeM` date DEFAULT NULL,
  `NueNR` varchar(10) DEFAULT NULL,
  `NueNR1` varchar(10) DEFAULT NULL,
  `NueNR2` varchar(10) DEFAULT NULL,
  `NueNR3` varchar(10) DEFAULT NULL,
  `NueNR4` varchar(10) DEFAULT NULL,
  `NueEst` int(1) DEFAULT NULL,
  `UsuRes1` int(4) DEFAULT NULL,
  `ProEsRx` int(1) DEFAULT NULL,
  `ProEsLa` int(1) DEFAULT NULL,
  `ProEsMe` int(1) DEFAULT NULL,
  `UsuRes2` int(4) DEFAULT NULL,
  `NueObs` mediumtext,
  PRIMARY KEY (`NueId`)
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `usuarios`
-- ----------------------------
DROP TABLE IF EXISTS `usuarios`;
CREATE TABLE `usuarios` (
  `Usu_Id` int(3) NOT NULL AUTO_INCREMENT,
  `Usu_Nom` varchar(250) DEFAULT NULL,
  `Usu_Usu` varchar(50) DEFAULT NULL,
  `Usu_Cla` varchar(10) DEFAULT NULL,
  `Usu_Acc` varchar(15) DEFAULT NULL,
  `Usu_Est` int(1) DEFAULT NULL,
  PRIMARY KEY (`Usu_Id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Records 
-- ----------------------------
INSERT INTO `empresas` VALUES ('1','A.L.T.'), ('2','AANAC'), ('3','ACERGAL SRL.'), ('4','ADEMAJ'), ('5','AECID'), ('6','AFP FUTURO'), ('7','AFP PREVISION'), ('8','AGEN.NAL.HIDROCARBUROS'), ('9','AJINOMOTO DEL PERU'), ('10','ALANOCA LTDA.'), ('11','ALDA INVERS SRL.'), ('12','ALTIKNITS'), ('13','ALUVITEM SRL'), ('14','ALVARADO'), ('15','AMERICAN IRIS S. A.'), ('16','ANTENA UNO CANAL 6 S.R.L.'), ('17','ARMUS LTDA.'), ('18','ASCARRUNZ PAREDES E.'), ('19','ASEO LA PAZ LIMP'), ('20','ASOC  SERV  FINANCIEROS CAFETALEROS'), ('21','AUTORIZACION ATENCION SUCRE'), ('22','AVICOLA CARGER '), ('23','B.M QUTSOURCING SRL'), ('24','BARRET LUJAN ORLANDO'), ('25','BEICRUZ'), ('26','BELLCOS'), ('27','BM QUTSOURCING SRL'), ('28','BOCA NEGRA '), ('29','BOLIVIA T.V'), ('30','BOSQUE SUR SRL'), ('31','C.B.N.'), ('32','CAJA PETROLERA DE SALUD'), ('33','CALLE ROSOS JUAN'), ('34','CENTRO DE CAPACITACION EFG S.R.'), ('35','CIDRE'), ('36','CIES'), ('37','COBOFAR S.A.'), ('38','COLGATE'), ('39','COMIBOL'), ('40','CONDORI ZENTENO'), ('41','CONSULTORA AUDICRACK'), ('42','CRISTO AUTOGAS SRL'), ('43','CZETA'), ('44','DEINCO'), ('45','DICSA'), ('46','DIPEX LTDA'), ('47','DON BOSCO'), ('48','DUFRY'), ('49','E.P.S.A.S.'), ('50','EDITORIAL DON BOSCO'), ('51','EMBOL'), ('52','EMP. ASEO LA PAZ LIMPIA'), ('53','EMP. EST. MI TELEFERICO'), ('54','EMP. PUBL. SOCIAL AGUA Y SANEAMIENTO'), ('55','EMPACAR S.A.'), ('56','FABRICA DE FIDEOS'), ('57','FABRICAL'), ('58','FADES'), ('59','FARMACORP S. A'), ('60','FERMEDICAL S R L'), ('61','FERTIL DE LOS ANDES'), ('62','FLUICONS'), ('63','FOCAPICI'), ('64','FONDESIF'), ('65','FUN.INFOCAL LA PAZ'), ('66','FUND. BURGOS MARKA'), ('67','FUND. UNIV. SIMOS I.P.'), ('68','FUND.BURGOS MARCA'), ('69','FURND. MARIA Y AMALIA'), ('70','FUSIP'), ('71','GERIMEX'), ('72','GOB. MUNICIPAL HUMAMANTA'), ('73','H.R. SOLUTIONS'), ('74','HERDELBERG BOL. S.A.'), ('75','HOLA SRL '), ('76','HORMIPRET   '), ('77','HP MEDIACL'), ('78','I.E.L.B'), ('79','IBMETRO'), ('80','IBRO SRL.'), ('81','IDEPRO'), ('82','IMCRUZ'), ('83','IME'), ('84','IMP. EXP. LIBRAMAR'), ('85','INBOLPAK'), ('86','INFOCAL'), ('87','INM. ZURIEL S.R.L.'), ('88','INST BOL DE NEFROLOGIA '), ('89','INVERSIONES SUCRE-ISSA'), ('90','ISSA '), ('91','KANTUTANI'), ('92','KOLLPING'), ('93','LAB. EUROFARMA '), ('94','LAMBOL'), ('95','LOS CACTUS DE LOS ANDES'), ('96','MADISA'), ('97','MEGACENTER L.P.S.A.'), ('98','MIN COMUNICACION'), ('99','MONOPOL LTDA'), ('100','MOPETMEN SRL ');
INSERT INTO `empresas` VALUES ('101','MURURATA INVERSTEMENT'), ('102','NIBOL'), ('103','NUÑEZ DEL PRADO ASOCIADOS'), ('104','OCCIDENTAL BOLIVIA '), ('105','ORMADERA'), ('106','PEFORTE'), ('107','POTENZA S. A '), ('108','PROMISA S.A.'), ('109','PROSALUD'), ('110','QUIMBAYA'), ('111','QUIMICA INDUSTRIAL J. MONTES'), ('112','QUINOA FOODS SRL. '), ('113','RENTISTA '), ('114','RESIMIN BOL. SRL'), ('115','ROJAS TAMBO MARIO'), ('116','RTP'), ('117','RUAT '), ('118','SABSA'), ('119','SEINCO'), ('120','SENASAG'), ('121','SENSORIAL'), ('122','SERV LOGISTICOS J L'), ('123','SERV.DE AEROPUERTOS BOL.'), ('124','SEVERICHE ROSADO '), ('125','SOC. SAL. UNIV. SALESIANA'), ('126','STS'), ('127','SUTI SANA SRL'), ('128','T PROMOCIONA BOLIVIA'), ('129','TERSA S. A'), ('130','TOTALCETRUS'), ('131','T-PROMOCIONA'), ('132','TRANS PETROLERO BOL.'), ('133','UNIVERSIDAD TUPAK KATARI'), ('134','Y.P.F.B.'), ('135','ZENTENO CONDORI ');
INSERT INTO `impbol` VALUES ('8','TORREZ CONDORI CARLA','FEMENINO','1990-02-10','30','76254879','7854885','DIPEX LTDA','2018-02-01','2020-02-12','2020-02-15','2020-02-19','4569',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);
INSERT INTO `nuevos` VALUES ('7','FERNANDEZ RODRIGUEZ MARTINEZ','MASCULINO','1979-02-10','41','70565487','7625487','AFP FUTURO','2019-07-08','2020-02-12','2020-02-15','2020-02-18','123456',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL), ('8','TORREZ CONDORI CARLA','FEMENINO','1990-02-10','30','76254879','7854885','DIPEX LTDA','2018-02-01','2020-02-12','2020-02-15','2020-02-19','4569',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL);
