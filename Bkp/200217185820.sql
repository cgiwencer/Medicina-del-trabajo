/*
MySQL Backup
Source Server Version: 5.1.30
Source Database: medtrabajo
Date: 17/02/2020 18:58:20
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
) ENGINE=InnoDB AUTO_INCREMENT=137 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `fichas`
-- ----------------------------
DROP TABLE IF EXISTS `fichas`;
CREATE TABLE `fichas` (
  `FicId` int(2) NOT NULL AUTO_INCREMENT,
  `FicCol` varchar(15) DEFAULT NULL,
  `FicNum` int(2) DEFAULT NULL,
  PRIMARY KEY (`FicId`)
) ENGINE=InnoDB AUTO_INCREMENT=76 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `impbol`
-- ----------------------------
DROP TABLE IF EXISTS `impbol`;
CREATE TABLE `impbol` (
  `NueId` int(6) DEFAULT NULL,
  `NueTip` varchar(30) DEFAULT NULL,
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
  `NueRec` varchar(10) DEFAULT NULL,
  `NueSigep` varchar(2) DEFAULT NULL,
  `NueCobro` varchar(2) DEFAULT NULL,
  `NueEmb` int(2) DEFAULT NULL,
  `NueNR1` varchar(10) DEFAULT NULL,
  `NueNR2` varchar(10) DEFAULT NULL,
  `NueNR3` varchar(10) DEFAULT NULL,
  `NueNR4` varchar(10) DEFAULT NULL,
  `NueEst` int(1) DEFAULT NULL,
  `UsuRes1` varchar(30) DEFAULT NULL,
  `ProEsRx` int(1) DEFAULT NULL,
  `ProEsLa` int(1) DEFAULT NULL,
  `ProEsMe` int(1) DEFAULT NULL,
  `UsuRes2` varchar(30) DEFAULT NULL,
  `NueObs` mediumtext
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `medicos`
-- ----------------------------
DROP TABLE IF EXISTS `medicos`;
CREATE TABLE `medicos` (
  `MedId` int(2) NOT NULL AUTO_INCREMENT,
  `MedNom` varchar(200) DEFAULT NULL,
  `MedCon` varchar(10) DEFAULT NULL,
  `MedCol` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`MedId`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Table structure for `nuevos`
-- ----------------------------
DROP TABLE IF EXISTS `nuevos`;
CREATE TABLE `nuevos` (
  `NueId` int(6) NOT NULL AUTO_INCREMENT,
  `NueTip` varchar(30) DEFAULT NULL,
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
  `NueRec` varchar(10) DEFAULT NULL,
  `NueSigep` varchar(10) DEFAULT NULL,
  `NueCobro` varchar(10) DEFAULT NULL,
  `NueEmb` int(2) DEFAULT NULL,
  `NueNR1` varchar(10) DEFAULT NULL,
  `NueNR2` varchar(10) DEFAULT NULL,
  `NueNR3` varchar(10) DEFAULT NULL,
  `NueNR4` varchar(10) DEFAULT NULL,
  `NueEst` int(1) DEFAULT NULL,
  `UsuRes1` varchar(30) DEFAULT NULL,
  `ProEsRx` int(1) DEFAULT NULL,
  `ProEsLa` int(1) DEFAULT NULL,
  `ProEsMe` int(1) DEFAULT NULL,
  `UsuRes2` varchar(30) DEFAULT NULL,
  `NueObs` mediumtext,
  PRIMARY KEY (`NueId`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;

-- ----------------------------
--  Records 
-- ----------------------------
INSERT INTO `empresas` VALUES ('1','A.L.T.'), ('2','AANAC'), ('3','ACERGAL SRL.'), ('4','ADEMAJ'), ('5','AECID'), ('6','AFP FUTURO'), ('7','AFP PREVISION'), ('8','AGEN.NAL.HIDROCARBUROS'), ('9','AJINOMOTO DEL PERU'), ('10','ALANOCA LTDA.'), ('11','ALDA INVERS SRL.'), ('12','ALTIKNITS'), ('13','ALUVITEM SRL'), ('14','ALVARADO'), ('15','AMERICAN IRIS S. A.'), ('16','ANTENA UNO CANAL 6 S.R.L.'), ('17','ARMUS LTDA.'), ('18','ASCARRUNZ PAREDES E.'), ('19','ASEO LA PAZ LIMP'), ('20','ASOC  SERV  FINANCIEROS CAFETALEROS'), ('21','AUTORIZACION ATENCION SUCRE'), ('22','AVICOLA CARGER '), ('23','B.M QUTSOURCING SRL'), ('24','BARRET LUJAN ORLANDO'), ('25','BEICRUZ'), ('26','BELLCOS'), ('27','BM QUTSOURCING SRL'), ('28','BOCA NEGRA '), ('29','BOLIVIA T.V'), ('30','BOSQUE SUR SRL'), ('31','C.B.N.'), ('32','CAJA PETROLERA DE SALUD'), ('33','CALLE ROSOS JUAN'), ('34','CENTRO DE CAPACITACION EFG S.R.'), ('35','CIDRE'), ('36','CIES'), ('37','COBOFAR S.A.'), ('38','COLGATE'), ('39','COMIBOL'), ('40','CONDORI ZENTENO'), ('41','CONSULTORA AUDICRACK'), ('42','CRISTO AUTOGAS SRL'), ('43','CZETA'), ('44','DEINCO'), ('45','DICSA'), ('46','DIPEX LTDA'), ('47','DON BOSCO'), ('48','DUFRY'), ('49','E.P.S.A.S.'), ('50','EDITORIAL DON BOSCO'), ('51','EMBOL'), ('52','EMP. ASEO LA PAZ LIMPIA'), ('53','EMP. EST. MI TELEFERICO'), ('54','EMP. PUBL. SOCIAL AGUA Y SANEAMIENTO'), ('55','EMPACAR S.A.'), ('56','FABRICA DE FIDEOS'), ('57','FABRICAL'), ('58','FADES'), ('59','FARMACORP S. A'), ('60','FERMEDICAL S R L'), ('61','FERTIL DE LOS ANDES'), ('62','FLUICONS'), ('63','FOCAPICI'), ('64','FONDESIF'), ('65','FUN.INFOCAL LA PAZ'), ('66','FUND. BURGOS MARKA'), ('67','FUND. UNIV. SIMOS I.P.'), ('68','FUND.BURGOS MARCA'), ('69','FURND. MARIA Y AMALIA'), ('70','FUSIP'), ('71','GERIMEX'), ('72','GOB. MUNICIPAL HUMAMANTA'), ('73','H.R. SOLUTIONS'), ('74','HERDELBERG BOL. S.A.'), ('75','HOLA SRL '), ('76','HORMIPRET   '), ('77','HP MEDIACL'), ('78','I.E.L.B'), ('79','IBMETRO'), ('80','IBRO SRL.'), ('81','IDEPRO'), ('82','IMCRUZ'), ('83','IME'), ('84','IMP. EXP. LIBRAMAR'), ('85','INBOLPAK'), ('86','INFOCAL'), ('87','INM. ZURIEL S.R.L.'), ('88','INST BOL DE NEFROLOGIA '), ('89','INVERSIONES SUCRE-ISSA'), ('90','ISSA '), ('91','KANTUTANI'), ('92','KOLLPING'), ('93','LAB. EUROFARMA '), ('94','LAMBOL'), ('95','LOS CACTUS DE LOS ANDES'), ('96','MADISA'), ('97','MEGACENTER L.P.S.A.'), ('98','MIN COMUNICACION'), ('99','MONOPOL LTDA'), ('100','MOPETMEN SRL ');
INSERT INTO `empresas` VALUES ('101','MURURATA INVERSTEMENT'), ('102','NIBOL'), ('103','NUÑEZ DEL PRADO ASOCIADOS'), ('104','OCCIDENTAL BOLIVIA '), ('105','ORMADERA'), ('106','PEFORTE'), ('107','POTENZA S. A '), ('108','PROMISA S.A.'), ('109','PROSALUD'), ('110','QUIMBAYA'), ('111','QUIMICA INDUSTRIAL J. MONTES'), ('112','QUINOA FOODS SRL. '), ('113','RENTISTA '), ('114','RESIMIN BOL. SRL'), ('115','ROJAS TAMBO MARIO'), ('116','RTP'), ('117','RUAT '), ('118','SABSA'), ('119','SEINCO'), ('120','SENASAG'), ('121','SENSORIAL'), ('122','SERV LOGISTICOS J L'), ('123','SERV.DE AEROPUERTOS BOL.'), ('124','SEVERICHE ROSADO '), ('125','SOC. SAL. UNIV. SALESIANA'), ('126','STS'), ('127','SUTI SANA SRL'), ('128','T PROMOCIONA BOLIVIA'), ('129','TERSA S. A'), ('130','TOTALCETRUS'), ('131','T-PROMOCIONA'), ('132','TRANS PETROLERO BOL.'), ('133','UNIVERSIDAD TUPAK KATARI'), ('134','Y.P.F.B.'), ('135','ZENTENO CONDORI '), ('136','SERVIGEL');
INSERT INTO `fichas` VALUES ('1','ROJO','1'), ('2','ROJO','2'), ('3','ROJO','3'), ('4','ROJO','4'), ('5','ROJO','5'), ('6','ROJO','6'), ('7','ROJO','7'), ('8','ROJO','8'), ('9','ROJO','9'), ('10','ROJO','10'), ('11','ROJO','11'), ('12','ROJO','12'), ('13','ROJO','13'), ('14','ROJO','14'), ('15','ROJO','15'), ('16','VERDE','1'), ('17','VERDE','2'), ('18','VERDE','3'), ('19','VERDE','4'), ('20','VERDE','5'), ('21','VERDE','6'), ('22','VERDE','7'), ('23','VERDE','8'), ('24','VERDE','9'), ('25','VERDE','10'), ('26','VERDE','11'), ('27','VERDE','12'), ('28','VERDE','13'), ('29','VERDE','14'), ('30','VERDE','15'), ('31','NARANJA','1'), ('32','NARANJA','2'), ('33','NARANJA','3'), ('34','NARANJA','4'), ('35','NARANJA','5'), ('36','NARANJA','6'), ('37','NARANJA','7'), ('38','NARANJA','8'), ('39','NARANJA','9'), ('40','NARANJA','10'), ('41','NARANJA','11'), ('42','NARANJA','12'), ('43','NARANJA','13'), ('44','NARANJA','14'), ('45','NARANJA','15'), ('46','CELESTE','1'), ('47','CELESTE','2'), ('48','CELESTE','3'), ('49','CELESTE','4'), ('50','CELESTE','5'), ('51','CELESTE','6'), ('52','CELESTE','7'), ('53','CELESTE','8'), ('54','CELESTE','9'), ('55','CELESTE','10'), ('56','CELESTE','11'), ('57','CELESTE','12'), ('58','CELESTE','13'), ('59','CELESTE','14'), ('60','CELESTE','15'), ('61','ROSADO','1'), ('62','ROSADO','2'), ('63','ROSADO','3'), ('64','ROSADO','4'), ('65','ROSADO','5'), ('66','ROSADO','6'), ('67','ROSADO','7'), ('68','ROSADO','8'), ('69','ROSADO','9'), ('70','ROSADO','10'), ('71','ROSADO','11'), ('72','ROSADO','12'), ('73','ROSADO','13'), ('74','ROSADO','14'), ('75','ROSADO','15');
INSERT INTO `impbol` VALUES ('7',NULL,'FERNANDEZ RODRIGUEZ MARTINEZ','MASCULINO','1979-02-10','41','70565487','7625487','AFP FUTURO','2019-07-08','2020-02-12','2020-02-15','2020-02-18',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL), ('8',NULL,'TORREZ CONDORI  CARLA','FEMENINO','1990-02-10','30','76254879','7854885','DIPEX LTDA','2018-02-01','2020-02-12','2020-02-15','2020-02-19','','0','0','-1',NULL,NULL,NULL,NULL,'1',NULL,'0','-1',NULL,NULL,NULL), ('9',NULL,'COCA CARVALLO WILSON','MASCULINO','1976-06-16','44','77266990','','CAJA PETROLERA DE SALUD','2014-06-06','2020-02-13','2020-02-15','2020-02-20','','0','0','0',NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL), ('10',NULL,'PACHECO FLORES STEFANI','FEMENINO','1988-11-15','31','72571901','','CAJA PETROLERA DE SALUD','2005-02-10','2020-02-13','2020-02-17','2020-02-20','','0','-1',NULL,NULL,NULL,NULL,NULL,'3',NULL,'0','-1',NULL,NULL,NULL), ('11',NULL,'PEREZ PEREZ JUAN','MASCULINO','1980-05-15','40','70587458','','EMBOL','2018-02-15','2020-02-14','2020-02-20','2020-02-25','','0','0','0',NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL);
INSERT INTO `medicos` VALUES ('1','DRA. SUSANA TARQUINO','302','NARANJA'), ('2','DRA. CLAUDIA RIVERO','303','ROSADO'), ('3','DRA. ELICENY RONDON','304','ROJO'), ('4','DR. LUIS ILLANES','309','VERDE'), ('5','DR. FERNANDO IRIGOYEN','310','AMARILLO'), ('6','DR. RUDDY GISBERT','308','CELESTE');
INSERT INTO `nuevos` VALUES ('7',NULL,'FERNANDEZ RODRIGUEZ MARTINEZ','MASCULINO','1979-02-10','41','70565487','7625487','AFP FUTURO','2019-07-08','2020-02-12','2020-02-15','2020-02-18',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL), ('8',NULL,'TORREZ CONDORI  CARLA','FEMENINO','1990-02-10','30','76254879','7854885','DIPEX LTDA','2018-02-01','2020-02-12','2020-02-15','2020-02-19','','0','0','-1',NULL,NULL,NULL,NULL,'1',NULL,'0','-1',NULL,NULL,NULL), ('9',NULL,'COCA CARVALLO WILSON','MASCULINO','1976-06-16','44','77266990','','CAJA PETROLERA DE SALUD','2014-06-06','2020-02-13','2020-02-15','2020-02-20','','0','0','0',NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL), ('10',NULL,'PACHECO FLORES STEFANI','FEMENINO','1988-11-15','31','72571901','','CAJA PETROLERA DE SALUD','2005-02-10','2020-02-13','2020-02-17','2020-02-20','','0','-1',NULL,NULL,NULL,NULL,NULL,'3',NULL,'0','-1',NULL,NULL,NULL), ('11',NULL,'PEREZ PEREZ JUAN','MASCULINO','1980-05-15','40','70587458','','EMBOL','2018-02-15','2020-02-14','2020-02-20','2020-02-25','','0','0','0',NULL,NULL,NULL,NULL,'1',NULL,'-1','-1',NULL,NULL,NULL), ('12','PREOCUPACIONAL','RODRIGUEZ FERNANDEZ RODRIGO FERNANDO','MASCULINO','1980-11-01','39','77731000','60100007','SERVIGEL','2013-05-06','2020-02-17','2020-02-21','2020-02-25','','457','','0',NULL,NULL,NULL,NULL,'1','Carlos Giwencer',NULL,NULL,NULL,NULL,NULL);
INSERT INTO `usuarios` VALUES ('6','Carlos Giwencer','cgiwencer','cagisa','11111','1'), ('7','Rosa Mamani','rmamani','rosita','11111','1');
