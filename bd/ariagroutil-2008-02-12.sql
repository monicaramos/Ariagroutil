/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.45-community-nt : Database - ariagroutil
*********************************************************************
Server version : 5.0.45-community-nt
*/

SET NAMES utf8;

SET SQL_MODE='';

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `appmenus` */

CREATE TABLE `appmenus` (
  `aplicacion` varchar(15) default '0',
  `Name` varchar(100) default '0',
  `caption` varchar(100) default '0',
  `indice` tinyint(3) default '0',
  `padre` smallint(3) unsigned default '0',
  `orden` tinyint(3) unsigned default NULL,
  `Contador` smallint(5) unsigned default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `appmenus` */

/*Table structure for table `appmenususuario` */

CREATE TABLE `appmenususuario` (
  `aplicacion` varchar(15) NOT NULL default '0',
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `codigo` smallint(3) unsigned NOT NULL default '0',
  `tag` varchar(100) default '0',
  PRIMARY KEY  (`aplicacion`,`codusu`,`codigo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `appmenususuario` */

/*Table structure for table `avnic` */

CREATE TABLE `avnic` (
  `codavnic` int(6) unsigned NOT NULL,
  `nombrper` varchar(40) NOT NULL,
  `nifperso` varchar(9) default NULL,
  `nifrepre` varchar(9) default NULL,
  `codposta` varchar(5) NOT NULL,
  `nomcalle` varchar(26) NOT NULL,
  `poblacio` varchar(30) NOT NULL,
  `provinci` varchar(30) NOT NULL,
  `codialta` tinyint(1) NOT NULL,
  `codbanco` smallint(4) NOT NULL,
  `codsucur` smallint(4) NOT NULL,
  `cuentaba` varchar(10) NOT NULL,
  `digcontr` varchar(2) NOT NULL,
  `imporper` decimal(9,2) NOT NULL,
  `imporret` decimal(9,2) NOT NULL,
  `anoejerc` smallint(4) NOT NULL,
  `nifpers1` varchar(9) default NULL,
  `fechalta` date NOT NULL,
  `nombper1` varchar(40) default NULL,
  `nomcall1` varchar(26) default NULL,
  `poblaci1` varchar(30) default NULL,
  `provinc1` varchar(30) default NULL,
  `codpost1` varchar(6) default NULL,
  `fechavto` date NOT NULL,
  `porcinte` decimal(5,2) NOT NULL,
  `importes` decimal(9,2) NOT NULL,
  `codmacta` varchar(10) NOT NULL,
  `observac` varchar(250) default NULL,
  PRIMARY KEY  (`codavnic`,`anoejerc`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `avnic` */

insert into `avnic` (`codavnic`,`nombrper`,`nifperso`,`nifrepre`,`codposta`,`nomcalle`,`poblacio`,`provinci`,`codialta`,`codbanco`,`codsucur`,`cuentaba`,`digcontr`,`imporper`,`imporret`,`anoejerc`,`nifpers1`,`fechalta`,`nombper1`,`nomcall1`,`poblaci1`,`provinc1`,`codpost1`,`fechavto`,`porcinte`,`importes`,`codmacta`,`observac`) values (1,'PEPITO','22642805',NULL,'46007','3232','eff','vefef',1,33,1200,'123467890','10','0.00','0.00',2007,NULL,'2007-11-27',NULL,NULL,NULL,NULL,NULL,'2008-11-27','1.00','100.00','1170000000',NULL);

/*Table structure for table `cabfact` */

CREATE TABLE `cabfact` (
  `codsecci` smallint(2) unsigned NOT NULL,
  `letraser` varchar(1) NOT NULL,
  `numfactu` int(7) unsigned NOT NULL,
  `fecfactu` date NOT NULL,
  `ctaclien` varchar(10) NOT NULL,
  `observac` varchar(250) default NULL,
  `intconta` tinyint(4) NOT NULL,
  `baseiva1` decimal(8,2) default NULL,
  `baseiva2` decimal(8,2) default NULL,
  `baseiva3` decimal(8,2) default NULL,
  `impoiva1` decimal(8,2) default NULL,
  `impoiva2` decimal(8,2) default NULL,
  `impoiva3` decimal(8,2) default NULL,
  `imporec1` decimal(8,2) default NULL,
  `imporec2` decimal(8,2) default NULL,
  `imporec3` decimal(8,2) default NULL,
  `totalfac` decimal(8,2) default NULL,
  `tipoiva1` smallint(2) default NULL,
  `tipoiva2` smallint(2) default NULL,
  `tipoiva3` smallint(2) default NULL,
  `porciva1` decimal(4,2) default NULL,
  `porciva2` decimal(4,2) default NULL,
  `porciva3` decimal(4,2) default NULL,
  `codforpa` smallint(6) NOT NULL,
  `porcrec1` decimal(4,2) default NULL,
  `porcrec2` decimal(4,2) default NULL,
  `porcrec3` decimal(4,2) default NULL,
  PRIMARY KEY  (`codsecci`,`letraser`,`numfactu`,`fecfactu`),
  CONSTRAINT `cabfact_ibfk_1` FOREIGN KEY (`codsecci`) REFERENCES `seccion` (`codsecci`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `cabfact` */

insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000001,'2007-10-30','4100001','lalala estaba',1,'22.60','12.60',NULL,'0.90','2.02',NULL,NULL,NULL,NULL,'38.12',2,1,NULL,'4.00','16.00',NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000002,'2007-10-30','4100002',NULL,1,'15.25',NULL,NULL,'0.61',NULL,NULL,NULL,NULL,NULL,'15.86',2,NULL,NULL,'4.00',NULL,NULL,6,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000003,'2007-10-31','4100008',NULL,1,'130.60',NULL,NULL,'5.22',NULL,NULL,NULL,NULL,NULL,'135.82',2,NULL,NULL,'4.00',NULL,NULL,6,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000004,'2007-10-31','4100006','dasdasdas',1,'1500.00','1000.00',NULL,'60.00','160.00',NULL,NULL,NULL,NULL,'2720.00',2,1,NULL,'4.00','16.00',NULL,2,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000005,'2007-11-05','4100008',NULL,1,'1520.35',NULL,NULL,'60.81',NULL,NULL,NULL,NULL,NULL,'1581.16',2,NULL,NULL,'4.00',NULL,NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000006,'2007-11-05','4100005',NULL,1,'10.25',NULL,NULL,'0.41',NULL,NULL,NULL,NULL,NULL,'10.66',2,NULL,NULL,'4.00',NULL,NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000007,'2007-11-05','4100001',NULL,1,'20.62','11.25',NULL,'0.82','1.80',NULL,NULL,NULL,NULL,'34.49',2,1,NULL,'4.00','16.00',NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000008,'2007-11-05','4100001',NULL,1,'10.25',NULL,NULL,'0.41',NULL,NULL,NULL,NULL,NULL,'10.66',2,NULL,NULL,'4.00',NULL,NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000009,'2007-11-06','4100001',NULL,1,'1000.00',NULL,NULL,'160.00',NULL,NULL,NULL,NULL,NULL,'1160.00',1,NULL,NULL,'16.00',NULL,NULL,1,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000010,'2007-11-07','4100005','1111',0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,5,NULL,NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (1,'B',1000011,'2007-11-30','4100005',NULL,0,'20.25',NULL,NULL,'0.81',NULL,NULL,'0.61',NULL,NULL,'21.67',2,NULL,NULL,'4.00',NULL,NULL,3,'3.00',NULL,NULL);
insert into `cabfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`ctaclien`,`observac`,`intconta`,`baseiva1`,`baseiva2`,`baseiva3`,`impoiva1`,`impoiva2`,`impoiva3`,`imporec1`,`imporec2`,`imporec3`,`totalfac`,`tipoiva1`,`tipoiva2`,`tipoiva3`,`porciva1`,`porciva2`,`porciva3`,`codforpa`,`porcrec1`,`porcrec2`,`porcrec3`) values (4,'G',8000001,'2007-11-06','4300000010',NULL,0,'180.00','50.59',NULL,'7.20','8.09',NULL,'1.80','1.52',NULL,'249.20',1,3,NULL,'4.00','16.00',NULL,101,'1.00','3.00',NULL);

/*Table structure for table `concefact` */

CREATE TABLE `concefact` (
  `codconce` smallint(3) unsigned NOT NULL,
  `nomconce` varchar(30) NOT NULL,
  `codmacta` varchar(10) NOT NULL,
  `tipoiva` smallint(2) NOT NULL,
  `numconta` smallint(2) NOT NULL,
  PRIMARY KEY  (`codconce`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `concefact` */

insert into `concefact` (`codconce`,`nomconce`,`codmacta`,`tipoiva`,`numconta`) values (1,'CONCEPTO 1','7050001',2,2);
insert into `concefact` (`codconce`,`nomconce`,`codmacta`,`tipoiva`,`numconta`) values (2,'CONCEPTO 2','7090000000',3,1);
insert into `concefact` (`codconce`,`nomconce`,`codmacta`,`tipoiva`,`numconta`) values (3,'CONCPTO 3','7520001',1,2);
insert into `concefact` (`codconce`,`nomconce`,`codmacta`,`tipoiva`,`numconta`) values (4,'CONCEPTO 4','7000000006',1,1);

/*Table structure for table `linfact` */

CREATE TABLE `linfact` (
  `codsecci` smallint(2) unsigned NOT NULL,
  `letraser` varchar(1) NOT NULL,
  `numfactu` int(7) unsigned NOT NULL,
  `fecfactu` date NOT NULL,
  `numlinea` smallint(2) NOT NULL,
  `codconce` smallint(3) unsigned NOT NULL,
  `ampliaci` varchar(50) default NULL,
  `importe` decimal(8,2) NOT NULL,
  `tipoiva` smallint(2) NOT NULL,
  PRIMARY KEY  (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`),
  KEY `FK_linfact` (`codconce`),
  CONSTRAINT `linfact_ibfk_1` FOREIGN KEY (`codconce`) REFERENCES `concefact` (`codconce`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `linfact` */

insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000001,'2007-10-30',1,1,'LALAAL','12.35',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000001,'2007-10-30',2,3,'52525','2.35',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000001,'2007-10-30',3,3,'ESCALA MUSICAL','10.25',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000001,'2007-10-30',4,1,'LALA','10.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000002,'2007-10-30',1,1,'LINEA 1','15.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000003,'2007-10-31',1,1,'IMPRESION DE LINEA 2','120.35',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000003,'2007-10-31',2,1,'12345678901234567890123456789012345678901234567890','10.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000004,'2007-10-31',1,1,'DASDASDAS','1500.00',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000004,'2007-10-31',2,3,'HOIKLK','1000.00',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000005,'2007-11-05',1,1,'LINEA DE IMPORTE 1','1520.35',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000006,'2007-11-05',1,1,'25252','10.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000007,'2007-11-05',1,1,'LINEA 1','10.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000007,'2007-11-05',2,3,'LINEA 2','11.25',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000007,'2007-11-05',3,1,'LINEA 3','10.37',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000008,'2007-11-05',1,1,'LINE A','10.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000009,'2007-11-06',1,3,'FSFSFDS','1000.00',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (1,'B',1000011,'2007-11-30',1,1,'111111','20.25',2);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (4,'G',8000001,'2007-11-06',1,4,'LIMPIEZA DE ESCANER','70.00',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (4,'G',8000001,'2007-11-06',2,4,'MENSAJERIA','150.00',1);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (4,'G',8000001,'2007-11-06',3,2,'ESTA LINEA ES LA TERCERA','35.23',3);
insert into `linfact` (`codsecci`,`letraser`,`numfactu`,`fecfactu`,`numlinea`,`codconce`,`ampliaci`,`importe`,`tipoiva`) values (4,'G',8000001,'2007-11-06',4,2,'OTRA AL 16','15.36',3);

/*Table structure for table `movim` */

CREATE TABLE `movim` (
  `codavnic` int(6) unsigned NOT NULL,
  `anoejerc` smallint(4) unsigned NOT NULL,
  `fechamov` date NOT NULL,
  `concepto` varchar(35) NOT NULL,
  `timporte` decimal(10,2) default NULL,
  `intconta` tinyint(1) NOT NULL,
  `timport1` decimal(10,2) NOT NULL,
  `timport2` decimal(10,2) NOT NULL,
  PRIMARY KEY  (`codavnic`,`anoejerc`,`fechamov`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `movim` */

/*Table structure for table `pcs` */

CREATE TABLE `pcs` (
  `codpc` smallint(5) unsigned NOT NULL default '0',
  `nompc` char(30) default NULL,
  `Conta` smallint(5) unsigned default NULL,
  PRIMARY KEY  (`codpc`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `pcs` */

/*Table structure for table `scryst` */

CREATE TABLE `scryst` (
  `codcryst` smallint(4) unsigned NOT NULL default '0',
  `nomcryst` varchar(30) NOT NULL default '',
  `documrpt` varchar(100) NOT NULL default '',
  `codigiso` varchar(10) default NULL,
  `codigrev` tinyint(2) unsigned default NULL,
  `lineapi1` varchar(140) default NULL,
  `lineapi2` varchar(140) default NULL,
  `lineapi3` varchar(140) default NULL,
  `lineapi4` varchar(140) default NULL,
  `lineapi5` varchar(140) default NULL,
  PRIMARY KEY  (`codcryst`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Tipos de Documentos';

/*Data for the table `scryst` */

insert into `scryst` (`codcryst`,`nomcryst`,`documrpt`,`codigiso`,`codigrev`,`lineapi1`,`lineapi2`,`lineapi3`,`lineapi4`,`lineapi5`) values (1,'Facturas Varias','CatFactVarias.rpt','',NULL,'','','','','');

/*Table structure for table `seccion` */

CREATE TABLE `seccion` (
  `codsecci` smallint(2) unsigned NOT NULL,
  `nomsecci` varchar(30) NOT NULL,
  `contador` int(7) NOT NULL,
  `letraser` varchar(1) NOT NULL,
  `numconta` smallint(2) NOT NULL,
  `raizcta` varchar(10) NOT NULL,
  PRIMARY KEY  (`codsecci`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `seccion` */

insert into `seccion` (`codsecci`,`nomsecci`,`contador`,`letraser`,`numconta`,`raizcta`) values (1,'TRATAMIENTOS Y SERVICIOS',1000011,'B',2,'4100');
insert into `seccion` (`codsecci`,`nomsecci`,`contador`,`letraser`,`numconta`,`raizcta`) values (2,'SECCIONES',2000000,'A',1,'12110');
insert into `seccion` (`codsecci`,`nomsecci`,`contador`,`letraser`,`numconta`,`raizcta`) values (3,'TTOS Y SERVICIOS',4000000,'C',2,'1190');
insert into `seccion` (`codsecci`,`nomsecci`,`contador`,`letraser`,`numconta`,`raizcta`) values (4,'ORTORSEA',8000001,'G',1,'43000');
insert into `seccion` (`codsecci`,`nomsecci`,`contador`,`letraser`,`numconta`,`raizcta`) values (5,'OTROS',7000000,'E',1,'11702');

/*Table structure for table `seglinea` */

CREATE TABLE `seglinea` (
  `codlinea` smallint(3) unsigned NOT NULL,
  `nomlinea` varchar(25) NOT NULL,
  PRIMARY KEY  (`codlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `seglinea` */

insert into `seglinea` (`codlinea`,`nomlinea`) values (2,'AUTOMATICA');
insert into `seglinea` (`codlinea`,`nomlinea`) values (9,'AUTOMATICA');
insert into `seglinea` (`codlinea`,`nomlinea`) values (15,'AUTOMATICA');
insert into `seglinea` (`codlinea`,`nomlinea`) values (100,'AUTOMATICA');

/*Table structure for table `segpoliza` */

CREATE TABLE `segpoliza` (
  `codrefer` varchar(10) NOT NULL,
  `codiplan` smallint(4) unsigned NOT NULL,
  `codlinea` smallint(3) unsigned NOT NULL,
  `colectiv` int(8) unsigned NOT NULL,
  `codmacta` varchar(10) NOT NULL,
  `nifasegu` varchar(9) NOT NULL,
  `nomasegu` varchar(35) NOT NULL,
  `fechaenv` date NOT NULL,
  `imppoliz` decimal(8,2) NOT NULL,
  `impinter` decimal(8,2) default NULL,
  `impampli` decimal(8,2) default NULL,
  `impreduc` decimal(8,2) default NULL,
  `intconta` tinyint(1) NOT NULL default '0',
  `inttesor` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`codrefer`,`codiplan`,`codlinea`),
  KEY `FK_segpoliza` (`codlinea`),
  CONSTRAINT `segpoliza_ibfk_1` FOREIGN KEY (`codlinea`) REFERENCES `seglinea` (`codlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `segpoliza` */

insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313722P2',2006,100,13609956,'1150000000','20698873T','BONO CELDA,DOLORES','2007-12-13','53.61',NULL,'1.25','12.35',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313724P1',2006,100,13609956,'1170002','20699362Y','LLORENS PONS,JOSE','2006-12-20','57.87',NULL,'0.00','0.00',1,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313725P4',2006,100,13609956,'1960001','20746015S','BARBERA SABATER,JOSE','2006-12-20','394.60',NULL,'0.00','0.00',1,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313726P0',2006,100,13609956,'4007200604','20763945M','GIL BONO,VICENTE','2006-12-20','1998.12',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313727P3',2006,100,13609956,'4007201103','20810163Q','BARBERA BISBAL,JOSE ENRIQUE','2006-12-20','147.84',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313728P6',2006,100,13609956,'4007201104','73772548W','BARBERA BISBAL,ANTONIO','2006-12-20','164.14',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313729P2',2006,100,13609956,'4007202016','73932632Y','CLIMENT GONZALEZ,ANTONIO','2006-12-20','2040.46',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313861P6',2006,100,13609956,'4007200561','20698890V','BONO BAGUENA,VICENTE','2006-12-22','227.56',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313862P2',2006,100,13609956,'4007201185','20776483P','DOS SANTOS,AMARAL BEATRIZ','2006-12-22','257.54',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313863P5',2006,100,13609956,'4007201462','20778585V','ORTIZ SANTIAGO,ANTONIO','2006-12-22','2024.87',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313864P1',2006,100,13609956,'4007201128','20794653P','BONO ESTEVE,TRINITARIO','2006-12-22','311.17',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313865P4',2006,100,13609956,'4007201063','20795923J','SIMON BISBAL,Mª ANGELES','2006-12-22','397.88',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313867P3',2006,100,13609956,'4007202129','20847589K','ALVENTOSA GARCIA,ARANTCHA','2006-12-22','591.33',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313868P6',2006,100,13609956,'4007201882','44861983T','CLIMENT RUBIO,MAURICIO','2006-12-22','266.11',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313951P3',2006,9,13609934,'4007202064','73928932D','HERVAS HERVAS,MERCEDES','2006-12-27','884.13',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313952P6',2006,100,13609956,'4007201567','20745410P','BUADES ARTES,RIGOBERTO','2006-12-27','2469.77',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313953P2',2006,100,13609956,'4007200517','20746038S','CLIMENT BISBAL,VICENTE','2006-12-27','731.11',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313954P5',2006,100,13609956,'4007202141','20855008B','CLIMENT MARTINEZ,DANIEL','2006-12-27','224.71',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('313955P1',2006,100,13609956,'4007200840','73929701L','MARTINEZ LLORENS,ROSA','2006-12-27','497.89',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314020P0',2006,100,13609956,'4007201311','19484666D','ALVAREZ PERPIÑA,AMPARO','2006-12-28','166.07',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314021P3',2006,100,13609956,'4007200623','20697941B','PERIS FERRANDO,JOSE SALVADOR','2006-12-28','326.21',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314022P6',2006,100,13609956,'4007200835','20698545V','GIL JOARES,JOSEFA','2006-12-28','237.47',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314023P2',2006,100,13609956,'4007200347','20746042L','GIL ESTARLICH,JOSE','2006-12-28','1051.34',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314024P5',2006,100,13609956,'4007200773','20776281J','FLORDELIS APARISI,MARIA CARMEN','2006-12-28','239.58',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314025P1',2006,100,13609956,'4007200825','20784031N','PERIS GIL,JOSE SALVADOR','2006-12-28','733.61',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314026P4',2006,100,13609956,'4007202117','20799397Z','LLEDO TORRES,RAFAEL','2006-12-28','336.31',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314027P0',2006,100,13609956,'4007202077','20800802Q','RUBIO POLO,JOSE','2006-12-28','90.04',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314030P2',2006,100,13609956,'4007202145','73930421A','OCHEDA HERVAS,AMPARO','2006-12-28','805.72',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314117P4',2006,9,13609934,'4007200310','20745711X','LLORENS SANZ,JOSE LUIS','2006-12-29','772.11',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314118P0',2006,100,13609956,'4007200233','20698643T','OSCA SANANTON,JOSE','2006-12-29','269.13',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314119P3',2006,100,13609956,'4007200210','20713968F','BARBERA FLORDELIS,JOSE LUIS','2006-12-29','327.38',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314120P6',2006,100,13609956,'4007200263','20738639E','CLIMENT BARBERA,JOSE ENRIQUE','2006-12-29','589.08',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314121P2',2006,100,13609956,'4007201256','20743887A','MARTINEZ GARCIA,CARMEN','2006-12-29','1078.02',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314122P5',2006,100,13609956,'4007200864','20746010X','CLIMENT BONO,VICENTA','2006-12-29','241.30',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314123P1',2006,100,13609956,'4007202146','20746048W','BISBAL BISBAL,JOSEFA','2006-12-29','137.35',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314124P4',2006,100,13609956,'4007201275','20781650T','PERERA YAGUES,VICENTE','2006-12-29','413.94',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314125P0',2006,100,13609956,'4007201269','20787515T','BARBERA MARTINEZ,JOSE LUIS','2006-12-29','472.89',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314126P3',2006,100,13609956,'4007200928','20790666T','CLIMENT RUBIO,MARIA CONSUELO','2006-12-29','589.86',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314127P6',2006,100,13609956,'4007201045','20807175H','SABATER BISBAL,JUAN MANUEL','2006-12-29','234.56',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314128P2',2006,100,13609956,'4007200160','36805656K','GIMENO LOPEZ,FERNANDO','2006-12-29','239.58',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314130P1',2006,100,13609956,'4007201279','73926444M','MONTAGUD ORTIZ,JOSEFA','2006-12-29','382.98',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314131P4',2006,100,13609956,'4007200588','73930892Z','MARTINEZ DIEZ,MODESTO','2006-12-29','143.31',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314180P4',2006,100,13609956,'4007200809','20726929L','SANZ AÑO,JOSEFA','2007-01-02','412.06',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314182P3',2006,100,13609956,'4007200158','20755455W','BISBAL NOVERQUES,VICENTE','2007-01-02','676.45',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314183P6',2006,100,13609956,'4007200853','20779725F','LOPEZ AÑO,JOSE','2007-01-02','281.86',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314184P2',2006,100,13609956,'4007202142','20787231S','CLIMENT SERRANO,MARIA CONSUELO','2007-01-02','207.91',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314185P5',2006,100,13609956,'4007200870','20792812F','ORTIZ SANTIAGO,VICENTE','2007-01-02','799.87',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314186P1',2006,100,13609956,'4007200812','20792930X','MOLLA CLIMENT,ANTONIO MANUEL','2007-01-02','647.54',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314187P4',2006,100,13609956,'4007202140','24382936S','ALVENTOSA GARCIA,OSCAR','2007-01-02','662.79',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314188P0',2006,100,13609956,'4007200634','73929718J','CLIMENT TOLMOS,JUAN LUIS','2007-01-02','775.90',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314189P3',2006,100,13609956,'4007200731','73942712N','ASENSI BONO,EMILIO','2007-01-02','197.30',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314220P5',2006,100,13609956,'4007200747','19343808A','SANZ LLORENS,JUANA','2007-01-03','174.04',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314221P1',2006,100,13609956,'4007201393','20746112C','LUCAS BONO,BERNARDO','2007-01-03','560.81',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314222P4',2006,100,13609956,'4007200371','20762138S','CLIMENT APARISI,VICENTE ENRIQUE','2007-01-03','213.25',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314223P0',2006,100,13609956,'4007200639','20763498H','GIL BONO,AMADOR','2007-01-03','567.55',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314224P3',2006,100,13609956,'4007201014','20770520W','BARBERA MARTINEZ,MIGUEL ANGEL','2007-01-03','169.12',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314225P6',2006,100,13609956,'4007200654','20780763X','GARCIA DIRANZO,ANA JUANA','2007-01-03','243.71',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314226P2',2006,100,13609956,'4007201942','20815988E','BONO DOSANTOS,JOSE ANTONIO','2007-01-03','425.39',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314227P5',2006,100,13609956,'4007202047','20833974E','CLIMENT BARBERA,VICENTE','2007-01-03','319.32',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314228P1',2006,100,13609956,'4007202048','20833975T','CLIMENT BARBERA,MARIA AMPARO','2007-01-03','136.78',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314229P4',2006,100,13609956,'4007202043','20834682V','ESCAMILLA ORTIZ,JOSE','2007-01-03','155.94',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314230P0',2006,100,13609956,'4007202132','20847204G','ALVENTOSA GARCIA,SANDRA','2007-01-03','421.29',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314231P3',2006,100,13609956,'4007201530','22620591E','BARBERA SAEZ,DESAMPARADOS','2007-01-03','140.92',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314232P6',2006,100,13609956,'4007201730','73549618B','BONO DONAT,VERONICA','2007-01-03','295.67',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314233P2',2006,100,13609956,'4007200614','73925136P','SANANTON ESQUER,VICENTE','2007-01-03','364.42',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314267P6',2006,100,13609956,'4007202062','19596033X','LAFUENTE PEÑARROCHA,CARMEN','2007-01-04','152.32',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314268P2',2006,100,13609956,'4007200010','19759075M','CLIMENT BARBERA,ENRIQUE','2007-01-04','112.74',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314269P5',2006,100,13609956,'4007200889','20698266Z','BONO ORTIZ,CONCEPCION','2007-01-04','480.74',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314271P4',2006,100,13609956,'4007200504','20755515Q','CLIMENT APARISI,RAFAEL','2007-01-04','207.91',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314272P0',2006,100,13609956,'4007200681','20765722B','ADAM BOLUDA,CARLOS','2007-01-04','410.35',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314273P3',2006,100,13609956,'4007200738','20778095X','BISBAL MOLLA,MARIA DOLORES','2007-01-04','277.00',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314274P6',2006,100,13609956,'4007200880','20778455W','LLISO SOLA,JOSE','2007-01-04','97.48',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314275P2',2006,100,13609956,'4007201312','20779726P','MINGUEZ SANZ,FRANCISCO','2007-01-04','24.37',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314276P5',2006,100,13609956,'4007200777','20792856M','CLIMENT BONO,VICENTE ENRIQUE','2007-01-04','592.16',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314277P1',2006,100,13609956,'4007200833','20807205W','CLIMENT BONO,JOSE CARLOS','2007-01-04','507.34',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314278P4',2006,100,13609956,'4007200946','25422711Y','CLIMENT APARISI,RAFAEL','2007-01-04','535.53',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314279P0',2006,100,13609956,'4007201560','25422792H','CLIMENT APARISI,MARIA ROSARIO','2007-01-04','94.98',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314282P2',2006,100,13609956,'4007202049','73772547R','BARBERA BISBAL,MARIA JESUS','2007-01-04','90.43',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314283P5',2006,100,13609956,'4007200390','73924974F','BISBAL ESTRUCH,AURELIO','2007-01-04','121.85',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314284P1',2006,100,13609956,'4007202144','73935218Q','SANZ BALLESTER,ANA','2007-01-04','212.04',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314301P3',2006,9,13609934,'4007200698','20699294F','ADAM CANET,CONCEPCION','2007-01-05','594.02',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314302P6',2006,9,13609934,'4007201843','20699356T','FORES BALLESTER,DOLORES','2007-01-05','643.23',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314303P2',2006,9,13609934,'4007201339','20699393Z','NOVERQUES FLORDELIS,PEDRO JUAN','2007-01-05','728.68',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314326P1',2006,9,13609934,'4007200723','20755638R','SALVADOR BISBAL,JOSE','2007-01-08','463.15',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314328P0',2006,100,13609956,'4007200569','19802952K','BISBAL ESCAMILLA,JOAQUIN','2007-01-08','213.24',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314330P6',2006,100,13609956,'4007200693','20727709V','BARBERA CLIMENT,ANTONIO','2007-01-08','182.77',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314331P2',2006,100,13609956,'4007200803','20745970Q','CLIMENT APARISI,REMEDIOS','2007-01-08','466.27',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314332P5',2006,100,13609956,'4007200600','20754606G','CABEZON MECINAS,JUAN','2007-01-08','197.30',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314333P1',2006,100,13609956,'4007201102','20771568S','BISBAL ROIG,VICENTE ENRIQUE','2007-01-08','164.14',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314392P3',2006,9,13609934,'4007202080','73549621Z','GIL BARBERA,MARIA JOSE','2007-01-09','187.75',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314393P6',2006,9,13609934,'4007200779','73934342Z','GIL OSCA,DOLORES','2007-01-09','773.71',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314394P2',2006,15,13609945,'4007202080','73549621Z','GIL BARBERA,MARIA JOSE','2007-01-09','47.54',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314395P5',2006,100,13609956,'4007200593','19437525H','ROSA ROVIRA,ENRIQUE JAIME','2007-01-09','440.66',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314396P1',2006,100,13609956,'4007200896','20766913Y','GARCIA BALLESTER,JOSEFA','2007-01-09','365.56',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314397P4',2006,100,13609956,'4007200972','24340849H','DURA MIQUEL,BEGOÑA','2007-01-09','820.69',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314398P0',2006,100,13609956,'4007201781','73561825M','ROSA BISBAL,MARIA VICENTA','2007-01-09','474.92',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314399P3',2006,100,13609956,'4007201861','73772582J','CORBERA GIL,VICENTE MIGUEL','2007-01-09','410.35',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314400P6',2006,100,13609956,'4007202118','73929707W','FLORDELIS APARISI,CECILIA','2007-01-09','225.48',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314401P2',2006,100,13609956,'4007202065','73938809L','NAVARRO CHASAN,CARMEN','2007-01-09','206.03',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314456P6',2006,9,13609934,'4007200053','19370152N','BARBERA NOVERQUES,SEBASTIAN','2007-01-10','258.98',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314458P5',2006,9,13609934,'4007200332','20746054P','BISBAL ESTRUCH,JOSE MARIA','2007-01-10','1100.42',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314459P1',2006,9,13609934,'4007201968','73936711Z','CANET ESCAMILLA,MARIA CARMEN','2007-01-10','283.99',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314460P4',2006,100,13609956,'4007201974','20691705P','HERRAN HERVAS,VICENTA','2007-01-10','738.63',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314461P0',2006,100,13609956,'4007200700','20698542Z','LLISO BISBAL,ROSAURA','2007-01-10','490.20',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314462P3',2006,100,13609956,'4007201303','20699329L','ORTIZ BARBERA,MARIA','2007-01-10','346.59',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314463P6',2006,100,13609956,'4007200404','20722668J','DIEZ SOLA,SALVADOR','2007-01-10','348.10',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314464P2',2006,100,13609956,'4007200539','20746045E','ESTEVE ADAM,VICENTE','2007-01-10','168.91',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314465P5',2006,100,13609956,'4007200411','20746093R','LLORENS ANTON,VICENTE','2007-01-10','81.55',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314466P1',2006,100,13609956,'4007202143','20752039J','GIMENO GIL,ASUNCION','2007-01-10','334.31',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314467P4',2006,100,13609956,'4007202112','20769894C','MONZO HERRAN,ROSA MARIA','2007-01-10','189.96',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314468P0',2006,100,13609956,'4007201939','20769936Q','OSCA MATEU,JOSEFA','2007-01-10','472.71',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314469P3',2006,100,13609956,'4007200687','20784087E','LLORENS LLISO,PEDRO MANUEL','2007-01-10','1542.33',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314470P6',2006,100,13609956,'4007201114','73942546F','JUANES TORTOLA,JOSE MIGUEL','2007-01-10','599.91',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314517P0',2006,2,13609923,'4300000764','E96135439','TAPIA LOPEZ, C.B. ,','2007-01-11','165.72',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314518P3',2006,9,13609934,'4300000765','E96135439','TAPIA LOPEZ, C.B. ,','2007-01-11','1162.13',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314519P6',2006,9,13609934,'4300000766','20699308K','SANTIAGO BALLESTER,JOSE','2007-01-11','334.97',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314520P2',2006,9,13609934,'4300000767','20771629F','CELDA TARIN,EDUARDO','2007-01-11','2740.45',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314521P5',2006,15,13609945,'4300000768','73640261B','LAZA SOLA,FRANCISCO','2007-01-11','89.15',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314522P1',2006,100,13609956,'4300000769','20698287N','LLORENS SANZ,CELESTINO','2007-01-11','221.40',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314524P0',2006,100,13609956,'4300000770','20755516V','FERRANDO BISBAL,JOSE MARIA','2007-01-11','153.53',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314525P3',2006,100,13609956,'4300000771','20757130K','DONATE GIL,ALFREDO','2007-01-11','118.27',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314526P6',2006,100,13609956,'4300000773','20774565E','MORANTE VALERO,EDUARDO MATIAS','2007-01-11','111.97',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314527P2',2006,100,13609956,'4300000774','20775998Y','CANET NOVERQUES,MARIA REMEDIOS','2007-01-11','225.48',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314528P5',2006,100,13609956,'4300000775','20796169Y','SANTOTOMAS MARTINEZ,JOSE VICENTE','2007-01-11','755.44',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314530P4',2006,100,13609956,'4300000776','20836973P','FERRANDO CANET,JORGE','2007-01-11','765.99',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314531P0',2006,100,13609956,'4300000777','20836975X','FERRANDO CANET,DAVID','2007-01-11','281.86',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314532P3',2006,100,13609956,'4300000779','22527434S','SANZ BALLESTER,VICENTE','2007-01-11','296.60',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314533P6',2006,100,13609956,'4300000781','73570059M','FERRANDO CANET,JOSEP MARIA','2007-01-11','656.56',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314534P2',2006,100,13609956,'4300000784','73927727T','SOLA BONO,JOSE MARIA','2007-01-11','328.28',NULL,'0.00','0.00',0,1);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314582P6',2006,9,13609934,'4007201546','20737534K','CLARIANA TAMPIA,JOAQUIN','2007-01-12','1252.53',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314583P2',2006,100,13609956,'4007201162','20786795Q','BONO ESTEVE,MIGUEL ANGEL','2007-01-12','221.63',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314584P5',2006,100,13609956,'4007201118','20807186Y','CLIMENT ALVAREZ,ROBERTO','2007-01-12','260.46',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314585P1',2006,100,13609956,'4007202116','52641480T','GUIJARRO NATIVIDAD,JOSEFA','2007-01-12','183.20',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314586P4',2006,100,13609956,'4007202155','52743372W','CAMPOS BOSCH,M. AMPARO','2007-01-12','277.96',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314587P0',2006,100,13609956,'4007201610','73549615P','GIL ALFONSO,ANA PILAR','2007-01-12','177.40',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314588P3',2006,100,13609956,'4007200559','73933976Q','ALVENTOSA MARTINEZ,MIGUEL','2007-01-12','219.33',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314644P3',2006,100,13609956,'4007201292','19716837H','CHOLVI GIRONA,JOSEFA','2007-01-15','267.91',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314645P6',2006,100,13609956,'4007200653','20746027G','SANANTON ESQUER,AMELIA','2007-01-15','128.57',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314646P2',2006,100,13609956,'4007202041','20746564N','ALEMANY MORENO,ADOLFO','2007-01-15','141.06',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314647P5',2006,100,13609956,'4007200357','20757302D','FRECHINA RUIZ,CRISTOBAL','2007-01-15','533.11',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314648P1',2006,100,13609956,'4007202058','20776958T','FERRER AÑO,LAURA','2007-01-15','58.37',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314715P6',2006,100,13609956,'4007200718','20698661H','GIL JOARES,MARIA','2007-01-16','228.13',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314716P2',2006,100,13609956,'4007201144','20773247S','BALAGUER SERRANO,MIGUEL ANGEL','2007-01-16','26.28',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314717P5',2006,100,13609956,'4007200826','20780977V','PERIS GIL,BARBARA MARIA','2007-01-16','214.21',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314718P1',2006,100,13609956,'4007202147','22601926X','GIMENEZ LENDINEZ,ANDRES','2007-01-16','666.57',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314719P4',2006,100,13609956,'4007201983','73772479W','GIL BARBERA,LOURDES','2007-01-16','121.85',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314782P4',2006,9,13609934,'4007200920','22548211T','SANZ BALLESTER,MARIA CARMEN','2007-01-17','479.33',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314783P0',2006,100,13609956,'4007200644','20772705W','BARBERA JOARES,PEDRO','2007-01-17','429.96',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314784P3',2006,100,13609956,'4007202026','20781793M','NOGUES MARTINEZ,JOSEFA','2007-01-17','767.66',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314785P6',2006,100,13609956,'4007201453','20804228S','MOLLA CLIMENT,MARIA ANGELES','2007-01-17','445.14',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314786P2',2006,100,13609956,'4007202119','53603595W','BARBERA PUCHALT,PEDRO JUAN','2007-01-17','473.07',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314787P5',2006,100,13609956,'4007202120','73534850D','PUCHALT VIDAL,AMPARO','2007-01-17','42.28',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314788P1',2006,100,13609956,'4007200875','73925733F','MOLL CABANES,ROSA','2007-01-17','281.86',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314789P4',2006,100,13609956,'4007200876','73943837X','MIQUEL MOLL,ROSA','2007-01-17','140.43',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('314923P0',2006,100,13609956,'4007201507','A46352373','FUGARPA, S.A. ,','2007-01-18','1378.73',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('315042P0',2006,9,13609934,'4007200386','20699256S','ESTEVE CLIMENT,FRANCISCO','2007-01-19','225.76',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('315043P3',2006,9,13609934,'4007201576','73935957L','CABRERA HIGUERAS,JUAN','2007-01-19','1408.14',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('315044P6',2006,100,13609956,'4007202135','20764880C','SANZ BALLESTER,FIDEL','2007-01-19','155.41',NULL,'0.00','0.00',0,0);
insert into `segpoliza` (`codrefer`,`codiplan`,`codlinea`,`colectiv`,`codmacta`,`nifasegu`,`nomasegu`,`fechaenv`,`imppoliz`,`impinter`,`impampli`,`impreduc`,`intconta`,`inttesor`) values ('318038P0',2006,9,13609934,'4007201554','22513540J','MARTINEZ CERVERO,JUAN','2007-03-08','318.88',NULL,'0.00','0.00',0,0);

/*Table structure for table `sempre` */

CREATE TABLE `sempre` (
  `codempre` smallint(3) NOT NULL default '0',
  `nomempre` varchar(40) NOT NULL default '',
  `domempre` varchar(40) NOT NULL default '',
  `codposta` varchar(6) NOT NULL default '',
  `pobempre` varchar(35) NOT NULL default '',
  `proempre` varchar(35) NOT NULL default '',
  `nifempre` varchar(9) NOT NULL default '',
  `telempre` varchar(10) default NULL,
  `faxempre` varchar(10) default NULL,
  `wwwempre` varchar(100) default NULL,
  `maiempre` varchar(100) default NULL,
  `perempre` varchar(40) default NULL,
  PRIMARY KEY  (`codempre`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sempre` */

insert into `sempre` (`codempre`,`nomempre`,`domempre`,`codposta`,`pobempre`,`proempre`,`nifempre`,`telempre`,`faxempre`,`wwwempre`,`maiempre`,`perempre`) values (1,'AGRICOLA DE CATADAU S.COOP.V.','Plaza España, 1','46196','CATADAU','VALENCIA','F46038881','962550503','962990323',NULL,'conta.coopcatadau@coopcatadau.e.telefonica.net','Quique');

/*Table structure for table `sparam` */

CREATE TABLE `sparam` (
  `codparam` smallint(1) NOT NULL default '0',
  `serconta` varchar(20) default NULL,
  `usuconta` varchar(20) default NULL,
  `pasconta` varchar(20) default NULL,
  `numconta` smallint(2) default NULL,
  `ctagasto` varchar(10) default NULL COMMENT 'Cta.Contable contado',
  `ctareten` varchar(10) default NULL COMMENT 'Cta.Dif.Negativas.',
  `websoporte` varchar(100) default NULL,
  `diremail` varchar(50) default NULL,
  `smtphost` varchar(50) default NULL,
  `smtpuser` varchar(50) default NULL,
  `smtppass` varchar(50) default NULL,
  `porcinte` decimal(5,2) NOT NULL default '0.00',
  `porcrete` decimal(5,2) NOT NULL default '0.00',
  `concedebe` smallint(4) NOT NULL,
  `concehaber` smallint(4) NOT NULL,
  `numdiario` smallint(4) NOT NULL,
  `seguros` tinyint(1) NOT NULL default '0' COMMENT 'Hay seguros',
  `avnics` tinyint(1) NOT NULL default '0' COMMENT 'Hay Avnics',
  `telefonia` tinyint(1) NOT NULL default '0' COMMENT 'Hay Telefonia',
  `sercontaseg` varchar(20) default NULL COMMENT 'servidor conta Seguros',
  `usucontaseg` varchar(20) default NULL COMMENT 'usuario conta Seguros',
  `pascontaseg` varchar(20) default NULL COMMENT 'pwda conta Seguros',
  `numcontaseg` smallint(2) default NULL COMMENT 'numconta Seguros',
  `concedebeseg` smallint(4) NOT NULL default '0' COMMENT 'Conce Debe Seguros',
  `concehaberseg` smallint(4) NOT NULL default '0' COMMENT 'Conce Haber Seguros',
  `numdiarioseg` smallint(4) NOT NULL default '0' COMMENT 'NumDiario Seguros',
  `ctabancoseg` varchar(10) default NULL COMMENT 'CtaBanco Seguros',
  `raizctasocseg` varchar(10) default NULL COMMENT 'Raiz Cta Socios Seguros',
  `sercontatel` varchar(20) default NULL COMMENT 'servidor conta Telefonia',
  `usucontatel` varchar(20) default NULL COMMENT 'usuario conta Telefonia',
  `pascontatel` varchar(20) default NULL COMMENT 'pwd conta Telefonia',
  `numcontatel` smallint(2) default NULL COMMENT 'numconta Telefonia',
  `concedebetel` smallint(4) NOT NULL default '0' COMMENT 'Conce Debe Telefonia',
  `concehabertel` smallint(4) NOT NULL default '0' COMMENT 'Conce Haber Telefonia',
  `numdiariotel` smallint(4) NOT NULL default '0' COMMENT 'Numdiario Telefonia',
  `raizctasoctel` varchar(10) default NULL COMMENT 'Raiz Cta Socios Telefonia',
  `ctabancotel` varchar(10) default NULL COMMENT 'CtaBanco Telefonia',
  `ctaventatel` varchar(10) default NULL COMMENT 'CtaVenta Telefonia',
  `numserietel` varchar(1) NOT NULL,
  `sercontafac` varchar(20) default NULL,
  `usucontafac` varchar(20) default NULL,
  `pascontafac` varchar(20) default NULL,
  `facturasvarias` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`codparam`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sparam` */

insert into `sparam` (`codparam`,`serconta`,`usuconta`,`pasconta`,`numconta`,`ctagasto`,`ctareten`,`websoporte`,`diremail`,`smtphost`,`smtpuser`,`smtppass`,`porcinte`,`porcrete`,`concedebe`,`concehaber`,`numdiario`,`seguros`,`avnics`,`telefonia`,`sercontaseg`,`usucontaseg`,`pascontaseg`,`numcontaseg`,`concedebeseg`,`concehaberseg`,`numdiarioseg`,`ctabancoseg`,`raizctasocseg`,`sercontatel`,`usucontatel`,`pascontatel`,`numcontatel`,`concedebetel`,`concehabertel`,`numdiariotel`,`raizctasoctel`,`ctabancotel`,`ctaventatel`,`numserietel`,`sercontafac`,`usucontafac`,`pascontafac`,`facturasvarias`) values (1,'pcmoni','root','aritel',1,'1150000000','1200000000','www.ariadnasoftware.com','conta.coopcatadau@coopcatadau.e.telefonica.net','smtp.e.telefonica.net','conta.coopcatadau@coopcatadau.e.telefonica.net','00830962','4.00','18.00',1,2,1,1,1,1,'pcmoni','root','aritel',2,6,7,1,'5700002','1120','pcmoni','root','aritel',1,1,1,1,'43000','5720000001','7000000006','T','pcmoni','root','aritel',1);

/*Table structure for table `telmovil` */

CREATE TABLE `telmovil` (
  `numserie` varchar(1) NOT NULL,
  `numfactu` int(7) unsigned NOT NULL,
  `fecfactu` date NOT NULL,
  `codmacta` varchar(10) NOT NULL,
  `baseimpo` decimal(10,2) NOT NULL,
  `cuotaiva` decimal(10,2) NOT NULL,
  `totalfac` decimal(10,2) default NULL,
  `intconta` tinyint(1) NOT NULL,
  PRIMARY KEY  (`numserie`,`numfactu`,`fecfactu`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `telmovil` */

insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',5,'2005-08-22','4300001162','32.00','5.00','82.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',6,'2005-08-22','4300002017','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',7,'2005-08-22','4300002116','40.00','6.00','46.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',8,'2005-08-22','4300000604','112.00','17.00','130.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',10,'2005-09-21','4300001162','55.00','8.00','64.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',11,'2005-09-21','4300002017','42.00','6.00','49.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',12,'2005-09-21','4300002116','47.00','7.00','55.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',13,'2005-09-21','4300000604','91.00','14.00','106.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',15,'2005-10-24','4300001162','43.00','6.00','50.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',16,'2005-10-24','4300002017','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',17,'2005-10-24','4300002116','30.00','4.00','35.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',18,'2005-10-24','4300000604','151.00','24.00','175.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',20,'2005-11-11','4300001162','38.00','6.00','44.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',21,'2005-11-11','4300002017','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',22,'2005-11-11','4300002116','26.00','4.00','30.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',23,'2005-11-11','4300000604','134.00','21.00','155.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',25,'2005-11-11','4300000116','1.00','0.00','2.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',26,'2005-12-20','4300000212','3.00','0.00','4.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',27,'2005-12-20','4300001162','41.00','6.00','48.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',28,'2005-12-20','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',29,'2005-12-20','4300002116','25.00','4.00','29.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',30,'2005-12-20','4300000604','114.00','18.00','132.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',32,'2005-12-20','4300001730','1.00','0.00','2.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',33,'2005-12-20','4300000390','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',34,'2005-12-20','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',35,'2005-12-20','4300001000','5.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',36,'2006-01-11','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',37,'2006-01-11','4300001162','43.00','6.00','50.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',38,'2006-01-11','4300002017','25.00','4.00','29.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',39,'2006-01-11','4300002116','43.00','6.00','50.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',40,'2006-01-11','4300000604','128.00','20.00','148.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',42,'2006-01-11','4300000275','13.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',43,'2006-01-11','4300001730','33.00','5.00','38.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',44,'2006-01-11','4300000390','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',45,'2006-01-11','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',46,'2006-01-11','4300001000','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',111,'2008-02-12','1150000003','125.00','1.00','225.00',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',222,'2008-02-12','1150000003','125.00','1.00','225.00',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',232,'2007-06-20','1150000003','136.00','1.00','236.00',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60012,'2006-02-10','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60013,'2006-02-10','4300001162','41.00','6.00','48.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60014,'2006-02-10','4300002017','44.00','7.00','51.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60015,'2006-02-10','4300002116','27.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60016,'2006-02-10','4300000604','271.00','43.00','315.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60018,'2006-02-10','4300000275','27.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60019,'2006-02-10','4300001730','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60020,'2006-02-10','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60021,'2006-02-10','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60022,'2006-02-10','4300001000','13.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60023,'2006-03-20','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60024,'2006-03-20','4300001162','39.00','7.00','47.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60025,'2006-03-20','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60026,'2006-03-20','4300002116','19.00','3.00','23.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60027,'2006-03-20','4300000604','232.00','44.00','276.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60029,'2006-03-20','4300000275','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60030,'2006-03-20','4300001730','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60031,'2006-03-20','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60032,'2006-03-20','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60033,'2006-03-20','4300001000','11.00','2.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60034,'2006-03-20','4300000761','3.00','0.00','3.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60035,'2006-04-12','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60036,'2006-04-12','4300001162','41.00','6.00','47.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60037,'2006-04-12','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60038,'2006-04-12','4300002116','25.00','4.00','29.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60039,'2006-04-12','4300000604','259.00','41.00','300.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60040,'2006-04-12','4300000879','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60042,'2006-04-12','4300000275','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60043,'2006-04-12','4300001730','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60044,'2006-04-12','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60045,'2006-04-12','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60046,'2006-04-12','4300001000','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60047,'2006-04-12','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60065,'2006-05-23','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60066,'2006-05-23','4300001162','49.00','7.00','57.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60067,'2006-05-23','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60068,'2006-05-23','4300002116','29.00','4.00','33.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60069,'2006-05-23','4300000604','299.00','47.00','347.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60070,'2006-05-23','4300000692','3.00','0.00','4.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60071,'2006-05-23','4300000879','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60073,'2006-05-23','4300000275','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60074,'2006-05-23','4300001730','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60075,'2006-05-23','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60076,'2006-05-23','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60077,'2006-05-23','4300001000','20.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60078,'2006-05-23','4300000487','15.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60079,'2006-05-23','4300000487','12.00','1.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60080,'2006-05-23','4300000487','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60081,'2006-05-23','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60082,'2006-06-20','4300000212','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60083,'2006-06-20','4300001162','70.00','11.00','81.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60084,'2006-06-20','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60085,'2006-06-20','4300002116','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60086,'2006-06-20','4300000604','297.00','47.00','345.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60087,'2006-06-20','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60088,'2006-06-20','4300000879','15.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60089,'2006-06-20','4300002117','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60090,'2006-06-20','4300001311','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60091,'2006-06-20','4300001356','34.00','5.00','39.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60092,'2006-06-20','4300002126','6.00','0.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60093,'2006-06-20','4300001063','3.00','0.00','3.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60095,'2006-06-20','4300000275','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60096,'2006-06-20','4300001730','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60097,'2006-06-20','4300000390','10.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60098,'2006-06-20','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60099,'2006-06-20','4300001000','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60100,'2006-06-20','4300000487','20.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60101,'2006-06-20','4300000487','20.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60102,'2006-06-20','4300000487','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60103,'2006-06-20','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60104,'2006-07-27','4300000212','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60105,'2006-07-27','4300001162','48.00','7.00','56.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60106,'2006-07-27','4300002017','3.00','0.00','3.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60107,'2006-07-27','4300002116','26.00','4.00','30.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60108,'2006-07-27','4300000604','223.00','35.00','259.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60109,'2006-07-27','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60110,'2006-07-27','4300000879','18.00','2.00','21.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60111,'2006-07-27','4300002117','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60112,'2006-07-27','4300001311','6.00','0.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60113,'2006-07-27','4300001356','21.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60114,'2006-07-27','4300002126','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60115,'2006-07-27','4300001063','11.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60116,'2006-09-07','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60117,'2006-07-27','4300000275','33.00','5.00','38.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60118,'2006-07-27','4300001730','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60119,'2006-07-27','4300000119','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60120,'2006-07-27','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60121,'2006-07-27','4300001000','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60122,'2006-07-27','4300000487','21.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60123,'2006-07-27','4300000487','27.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60124,'2006-07-27','4300000487','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60125,'2006-07-27','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60126,'2006-07-27','4300000371','4.00','0.00','5.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60127,'2006-07-27','4300001918','5.00','0.00','5.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60128,'2006-07-27','4300000357','2.00','0.00','3.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60129,'2006-07-27','4300000078','1.00','0.00','1.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60130,'2006-07-27','4300000952','5.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60131,'2006-07-27','4300000588','3.00','0.00','3.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60132,'2006-09-07','4300000212','11.00','2.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60133,'2006-09-07','4300001162','39.00','7.00','47.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60134,'2006-09-07','4300002017','9.00','0.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60135,'2006-09-07','4300002116','30.00','5.00','36.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60136,'2006-09-07','4300000604','164.00','31.00','195.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60137,'2006-09-07','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60138,'2006-09-07','4300000879','15.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60139,'2006-09-07','4300002117','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60140,'2006-09-07','4300001311','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60141,'2006-09-07','4300001356','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60142,'2006-09-07','4300002126','14.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60143,'2006-09-07','4300001063','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60145,'2006-09-07','4300000275','13.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60146,'2006-09-07','4300001730','15.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60147,'2006-09-07','4300000390','16.00','3.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60149,'2006-09-07','4300001000','11.00','2.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60150,'2006-09-07','4300000487','16.00','3.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60151,'2006-09-07','4300000487','75.00','14.00','90.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60152,'2006-09-07','4300000487','11.00','2.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60153,'2006-09-07','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60154,'2006-09-07','4300000371','10.00','2.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60155,'2006-09-07','4300001918','28.00','5.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60156,'2006-09-07','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60157,'2006-09-07','4300000078','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60158,'2006-09-07','4300000952','17.00','3.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',60159,'2006-09-07','4300000588','11.00','2.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600035,'2007-02-14','4300001437','10.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600036,'2007-02-14','4300000212','14.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600037,'2007-02-14','4300001162','34.00','5.00','39.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600038,'2007-02-14','4300002116','32.00','5.00','38.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600039,'2007-02-14','4300000604','193.00','30.00','224.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600040,'2007-02-14','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600041,'2007-02-14','4300000879','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600042,'2007-02-14','4300002117','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600043,'2007-02-14','4300001311','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600044,'2007-02-14','4300001356','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600045,'2007-02-14','4300002126','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600046,'2007-02-14','4300001063','18.00','3.00','21.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600047,'2007-02-14','4300001331','42.00','6.00','49.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600048,'2007-02-14','4300000398','5.00','0.00','5.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600050,'2007-02-14','4300004500','91.00','14.00','105.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600051,'2007-02-14','4300000275','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600052,'2007-02-14','4300001730','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600053,'2007-02-14','4300000390','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600054,'2007-02-14','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600055,'2007-02-14','4300001000','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600056,'2007-02-14','4300000487','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600057,'2007-02-14','4300000487','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600058,'2007-02-14','4300000487','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600059,'2007-02-14','4300001143','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600060,'2007-02-14','4300001117','0.00','0.00','0.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600061,'2007-02-14','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600062,'2007-02-14','4300000979','23.00','3.00','27.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600063,'2007-02-14','4300000371','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600064,'2007-02-14','4300001918','27.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600065,'2007-02-14','4300002017','21.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600066,'2007-02-14','4300001393','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600067,'2007-02-14','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600068,'2007-02-14','4300000078','6.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600069,'2007-02-14','4300000952','26.00','4.00','30.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600070,'2007-02-14','4300000588','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600071,'2007-03-13','4300001497','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600072,'2007-03-13','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600073,'2007-03-13','4300001162','37.00','5.00','43.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600074,'2007-03-13','4300002116','35.00','5.00','41.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600075,'2007-03-13','4300000604','234.00','37.00','272.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600076,'2007-03-13','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600077,'2007-03-13','4300000879','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600078,'2007-03-13','4300002117','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600079,'2007-03-13','4300001311','8.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600080,'2007-03-13','4300001356','13.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600081,'2007-03-13','4300002126','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600082,'2007-03-13','4300001063','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600083,'2007-03-13','4300001331','24.00','3.00','27.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600084,'2007-03-13','4300000398','14.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600086,'2007-03-13','4300004500','54.00','8.00','62.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600087,'2007-03-13','4300000275','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600088,'2007-03-13','4300001730','33.00','5.00','39.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600089,'2007-03-13','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600090,'2007-03-13','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600091,'2007-03-13','4300001000','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600092,'2007-03-13','4300000487','12.00','1.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600093,'2007-03-13','4300000487','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600094,'2007-03-13','4300000487','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600095,'2007-03-13','4300001143','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600096,'2007-03-13','4300001117','20.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600097,'2007-03-13','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600098,'2007-03-13','4300000979','23.00','3.00','27.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600099,'2007-03-13','4300000371','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600100,'2007-03-13','4300001918','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600101,'2007-03-13','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600102,'2007-03-13','4300001393','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600103,'2007-03-13','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600104,'2007-03-13','4300000078','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600105,'2007-03-13','4300000952','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600106,'2007-03-13','4300000588','7.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600108,'2007-04-20','4300000212','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600109,'2007-04-20','4300001162','26.00','4.00','30.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600110,'2007-04-20','4300002116','26.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600111,'2007-04-20','4300000604','221.00','35.00','257.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600112,'2007-04-20','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600113,'2007-04-20','4300000879','15.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600114,'2007-04-20','4300002117','6.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600115,'2007-04-20','4300001311','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600116,'2007-04-20','4300001356','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600117,'2007-04-20','4300002126','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600118,'2007-04-20','4300001063','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600119,'2007-04-20','4300001331','26.00','4.00','31.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600120,'2007-04-20','4300000398','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600122,'2007-04-20','4300004500','60.00','9.00','70.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600123,'2007-04-20','4300000275','29.00','4.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600124,'2007-04-20','4300001730','34.00','5.00','40.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600125,'2007-04-20','4300000390','6.00','0.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600126,'2007-04-20','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600127,'2007-04-20','4300001000','18.00','2.00','21.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600128,'2007-04-20','4300000487','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600129,'2007-04-20','4300000487','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600130,'2007-04-20','4300000487','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600131,'2007-04-20','4300001143','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600132,'2007-04-20','4300002043','0.00','0.00','0.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600133,'2007-04-20','4300001117','28.00','4.00','33.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600134,'2007-04-20','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600135,'2007-04-20','4300000979','20.00','3.00','23.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600136,'2007-04-20','4300000371','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600137,'2007-04-20','4300001918','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600138,'2007-04-20','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600139,'2007-04-20','4300001393','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600140,'2007-04-20','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600141,'2007-04-20','4300000078','11.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600142,'2007-04-20','4300000952','11.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600143,'2007-04-20','4300000588','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600145,'2007-05-14','4300000212','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600146,'2007-05-14','4300001162','41.00','6.00','47.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600147,'2007-05-14','4300002116','29.00','4.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600148,'2007-05-14','4300000604','363.00','58.00','421.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600149,'2007-05-14','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600150,'2007-05-14','4300000879','15.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600151,'2007-05-14','4300002117','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600152,'2007-05-14','4300001311','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600153,'2007-05-14','4300001356','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600154,'2007-05-14','4300002126','21.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600155,'2007-05-14','4300001063','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600156,'2007-05-14','4300001331','79.00','12.00','91.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600157,'2007-05-14','4300000398','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600159,'2007-05-14','4300004500','84.00','13.00','98.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600160,'2006-09-27','4300000212','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600160,'2007-05-14','4300000275','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600161,'2006-09-27','4300001162','51.00','8.00','60.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600161,'2007-05-14','4300001730','12.00','1.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600162,'2006-09-27','4300002017','2.00','0.00','2.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600162,'2007-05-14','4300000390','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600163,'2006-09-27','4300002116','28.00','4.00','33.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600163,'2007-05-14','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600164,'2006-09-27','4300000604','283.00','45.00','329.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600164,'2007-05-14','4300001000','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600165,'2006-09-27','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600165,'2007-05-14','4300000487','22.00','3.00','25.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600166,'2006-09-27','4300000879','14.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600166,'2007-05-14','4300000487','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600167,'2006-09-27','4300002117','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600167,'2007-05-14','4300000487','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600168,'2006-09-27','4300001311','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600168,'2007-05-14','4300001143','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600169,'2006-09-27','4300001356','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600169,'2007-05-14','4300002043','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600170,'2006-09-27','4300002126','7.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600170,'2007-05-14','4300001117','11.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600171,'2006-09-27','4300001063','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600171,'2007-05-14','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600172,'2007-05-14','4300000979','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600173,'2006-09-27','4300000275','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600173,'2007-05-14','4300000371','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600174,'2006-09-27','4300001730','14.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600174,'2007-05-14','4300001918','22.00','3.00','26.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600175,'2006-09-27','4300000390','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600175,'2007-05-14','4300002017','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600176,'2006-09-27','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600176,'2007-05-14','4300001393','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600177,'2006-09-27','4300001000','7.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600177,'2007-05-14','4300000357','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600178,'2006-09-27','4300000487','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600178,'2007-05-14','4300000078','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600179,'2006-09-27','4300000487','87.00','14.00','101.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600179,'2007-05-14','4300000952','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600180,'2006-09-27','4300000487','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600180,'2007-05-14','4300000588','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600181,'2006-09-27','4300001254','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600182,'2006-09-27','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600183,'2006-09-27','4300000979','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600184,'2006-09-27','4300000371','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600185,'2006-09-27','4300001918','25.00','4.00','29.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600186,'2006-09-27','4300002017','17.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600187,'2006-09-27','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600188,'2006-09-27','4300001331','59.00','9.00','68.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600189,'2006-09-27','4300000078','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600190,'2006-09-27','4300000952','21.00','3.00','25.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600191,'2006-09-27','4300000588','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600193,'2006-11-08','4300000212','17.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600194,'2006-11-08','4300001162','47.00','7.00','54.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600195,'2006-11-08','4300002116','19.00','3.00','22.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600196,'2006-11-08','4300000604','232.00','37.00','269.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600197,'2006-11-08','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600198,'2006-11-08','4300000879','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600199,'2006-11-08','4300002117','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600200,'2006-11-08','4300001311','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600201,'2006-11-08','4300001356','29.00','4.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600202,'2006-11-08','4300002126','8.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600203,'2006-11-08','4300001063','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600204,'2006-11-08','4300001331','54.00','8.00','62.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600206,'2006-11-08','4300000275','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600207,'2006-11-08','4300001730','18.00','2.00','21.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600208,'2006-11-08','4300000390','15.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600209,'2006-11-08','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600210,'2006-11-08','4300001000','18.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600211,'2006-11-08','4300004500','28.00','4.00','32.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600212,'2006-11-08','4300000487','17.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600213,'2006-11-08','4300000487','52.00','8.00','61.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600214,'2006-11-08','4300000487','12.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600215,'2006-11-08','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600216,'2006-11-08','4300000979','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600217,'2006-11-08','4300000371','8.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600218,'2006-11-08','4300001918','24.00','3.00','28.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600219,'2006-11-08','4300002017','35.00','5.00','40.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600220,'2006-11-08','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600221,'2006-11-08','4300000078','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600222,'2006-11-08','4300000952','24.00','3.00','28.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600223,'2006-11-08','4300000588','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600225,'2006-11-20','4300000212','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600226,'2006-11-20','4300001162','37.00','6.00','43.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600227,'2006-11-20','4300002116','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600228,'2006-11-20','4300000604','312.00','49.00','362.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600229,'2006-11-20','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600230,'2006-11-20','4300000879','8.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600231,'2006-11-20','4300002117','11.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600232,'2006-11-20','4300001311','7.00','1.00','9.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600233,'2006-11-20','4300001356','29.00','4.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600234,'2006-11-20','4300002126','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600235,'2006-11-20','4300001063','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600236,'2006-11-20','4300001331','45.00','7.00','53.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600238,'2006-11-20','4300004500','65.00','10.00','75.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600239,'2006-11-20','4300000275','15.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600240,'2006-11-20','4300001730','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600241,'2006-11-20','4300000390','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600242,'2006-11-20','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600243,'2006-11-20','4300001000','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600244,'2006-11-20','4300000487','22.00','3.00','25.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600245,'2006-11-20','4300000487','23.00','3.00','27.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600246,'2006-11-20','4300000487','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600247,'2006-11-20','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600248,'2006-11-20','4300000979','15.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600249,'2006-11-20','4300000371','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600250,'2006-11-20','4300001918','24.00','3.00','28.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600251,'2006-11-20','4300002017','25.00','4.00','29.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600252,'2006-11-20','4300001393','29.00','4.00','34.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600253,'2006-11-20','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600254,'2006-11-20','4300000078','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600255,'2006-11-20','4300000952','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600256,'2006-11-20','4300000588','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600258,'2006-12-21','4300000212','14.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600259,'2006-12-21','4300001162','43.00','6.00','50.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600260,'2006-12-21','4300002116','23.00','3.00','27.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600261,'2006-12-21','4300000604','243.00','39.00','282.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600262,'2006-12-21','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600263,'2006-12-21','4300000879','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600264,'2006-12-21','4300002117','17.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600265,'2006-12-21','4300001311','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600266,'2006-12-21','4300001356','14.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600267,'2006-12-21','4300002126','13.00','2.00','16.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600268,'2006-12-21','4300001063','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600269,'2006-12-21','4300001331','33.00','5.00','38.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600271,'2006-12-21','4300004500','63.00','10.00','73.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600272,'2006-12-21','4300000275','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600273,'2006-12-21','4300001730','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600274,'2006-12-21','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600275,'2006-12-21','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600276,'2006-12-21','4300001000','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600277,'2006-12-21','4300000487','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600278,'2006-12-21','4300000487','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600279,'2006-12-21','4300000487','9.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600280,'2006-12-21','4300001143','4.00','0.00','4.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600281,'2006-12-21','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600282,'2006-12-21','4300000979','17.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600283,'2006-12-21','4300000371','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600284,'2006-12-21','4300001918','14.00','2.00','17.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600285,'2006-12-21','4300002017','24.00','3.00','28.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600286,'2006-12-21','4300001393','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600287,'2006-12-21','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600288,'2006-12-21','4300000078','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600289,'2006-12-21','4300000952','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',600290,'2006-12-21','4300000588','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700001,'2007-01-24','4300001437','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700002,'2007-01-24','4300000212','11.00','1.00','13.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700003,'2007-01-24','4300001162','42.00','6.00','48.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700004,'2007-01-24','4300002116','31.00','5.00','36.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700005,'2007-01-24','4300000604','412.00','65.00','478.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700006,'2007-01-24','4300000692','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700007,'2007-01-24','4300000879','9.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700008,'2007-01-24','4300002117','10.00','1.00','12.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700009,'2007-01-24','4300001311','8.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700010,'2007-01-24','4300001356','16.00','2.00','19.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700011,'2007-01-24','4300002126','21.00','3.00','25.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700012,'2007-01-24','4300001063','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700013,'2007-01-24','4300001331','51.00','8.00','59.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700015,'2007-01-24','4300004500','65.00','10.00','76.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700016,'2007-01-24','4300000275','20.00','3.00','24.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700017,'2007-01-24','4300001730','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700018,'2007-01-24','4300000390','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700019,'2007-01-24','4300000116','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700020,'2007-01-24','4300001000','12.00','2.00','14.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700021,'2007-01-24','4300000487','16.00','2.00','18.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700022,'2007-01-24','4300000487','18.00','3.00','21.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700023,'2007-01-24','4300000487','10.00','1.00','11.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700024,'2007-01-24','4300001143','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700025,'2007-01-24','4300000761','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700026,'2007-01-24','4300000979','24.00','4.00','28.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700027,'2007-01-24','4300000371','6.00','1.00','7.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700028,'2007-01-24','4300001918','23.00','3.00','26.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700029,'2007-01-24','4300002017','22.00','3.00','26.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700030,'2007-01-24','4300001393','13.00','2.00','15.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700031,'2007-01-24','4300000357','6.00','0.00','6.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700032,'2007-01-24','4300000078','7.00','1.00','8.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700033,'2007-01-24','4300000952','17.00','2.00','20.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',700034,'2007-01-24','4300000588','8.00','1.00','10.00',1);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900181,'2007-06-20','4300001497','8.16','1.31','9.47',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900182,'2007-06-20','4300000212','11.88','1.90','13.78',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900183,'2007-06-20','4300001162','43.47','6.96','50.43',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900184,'2007-06-20','4300002116','37.88','6.06','43.94',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900185,'2007-06-20','4300000604','189.36','30.30','219.66',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900186,'2007-06-20','4300000692','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900187,'2007-06-20','4300000879','18.81','3.01','21.82',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900188,'2007-06-20','4300002117','11.20','1.79','12.99',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900189,'2007-06-20','4300001311','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900190,'2007-06-20','4300001356','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900191,'2007-06-20','4300002126','32.21','5.15','37.36',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900192,'2007-06-20','4300001063','6.66','1.07','7.73',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900193,'2007-06-20','4300001331','54.07','8.65','62.72',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900194,'2007-06-20','4300000398','18.28','2.92','21.20',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900196,'2007-06-20','4300004500','128.42','20.55','148.97',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900197,'2007-06-20','4300000275','29.69','4.75','34.44',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900198,'2007-06-20','4300001730','8.55','1.37','9.92',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900199,'2007-06-20','4300002072','8.69','1.39','10.08',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900200,'2007-06-20','4300000390','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900201,'2007-06-20','4300000116','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900202,'2007-06-20','4300001000','15.10','2.42','17.52',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900203,'2007-06-20','4300000857','0.39','0.06','0.45',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900204,'2007-06-20','4300000487','21.36','3.42','24.78',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900205,'2007-06-20','4300000487','13.10','2.10','15.20',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900206,'2007-06-20','4300000487','10.41','1.67','12.08',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900207,'2007-06-20','4300001117','77.28','12.36','89.64',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900208,'2007-06-20','4300001143','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900209,'2007-06-20','4300002043','18.72','3.00','21.72',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900210,'2007-06-20','4300000761','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900211,'2007-06-20','4300000979','25.73','4.12','29.85',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900212,'2007-06-20','4300000371','15.65','2.50','18.15',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900213,'2007-06-20','4300001918','47.51','7.60','55.11',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900214,'2007-06-20','4300002017','10.11','1.62','11.73',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900215,'2007-06-20','4300001393','25.15','4.02','29.17',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900216,'2007-06-20','4300000357','7.57','1.21','8.78',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900217,'2007-06-20','4300000078','6.00','0.96','6.96',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900218,'2007-06-20','4300000952','19.73','3.16','22.89',0);
insert into `telmovil` (`numserie`,`numfactu`,`fecfactu`,`codmacta`,`baseimpo`,`cuotaiva`,`totalfac`,`intconta`) values ('T',900219,'2007-06-20','4300000588','6.26','1.00','7.26',0);

/*Table structure for table `tmpinformes` */

CREATE TABLE `tmpinformes` (
  `codusu` smallint(3) unsigned NOT NULL default '0',
  `codigo1` int(6) unsigned NOT NULL default '0',
  `fecha1` date default NULL,
  `fecha2` date default NULL,
  `campo1` smallint(4) unsigned default NULL,
  `campo2` smallint(4) unsigned default NULL,
  `nombre1` varchar(40) default NULL,
  `importe1` decimal(12,2) default NULL,
  `importe2` decimal(12,2) default NULL,
  `importe3` decimal(12,2) default NULL,
  `importe4` decimal(12,2) default NULL,
  `importe5` decimal(12,2) default NULL,
  `importe6` decimal(12,2) default NULL,
  `porcen1` decimal(5,2) default NULL,
  `porcen2` decimal(5,2) default NULL,
  `importeb1` decimal(12,2) default NULL,
  `importeb2` decimal(12,2) default NULL,
  `importeb3` decimal(12,2) default NULL,
  `importeb4` decimal(12,2) default NULL,
  `importeb5` decimal(12,2) default NULL,
  `importeb6` decimal(12,2) default NULL,
  `nombre2` varchar(40) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Temporal para informes';

/*Data for the table `tmpinformes` */

/*Table structure for table `usuarios` */

CREATE TABLE `usuarios` (
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `nomusu` char(30) NOT NULL default '',
  `dirfich` char(255) default NULL,
  `nivelusu` tinyint(1) NOT NULL default '-1',
  `login` char(20) NOT NULL default '',
  `passwordpropio` char(20) NOT NULL default '',
  `nivelusuges` tinyint(4) NOT NULL default '-1',
  PRIMARY KEY  (`codusu`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `usuarios` */

insert into `usuarios` (`codusu`,`nomusu`,`dirfich`,`nivelusu`,`login`,`passwordpropio`,`nivelusuges`) values (0,'root',NULL,0,'root','aritel',0);
insert into `usuarios` (`codusu`,`nomusu`,`dirfich`,`nivelusu`,`login`,`passwordpropio`,`nivelusuges`) values (1,'Mª Jose Rubio',NULL,-1,'mjose','mjose',1);
insert into `usuarios` (`codusu`,`nomusu`,`dirfich`,`nivelusu`,`login`,`passwordpropio`,`nivelusuges`) values (2,'Eva Llorens',NULL,-1,'eva','eva',1);

/*Table structure for table `zbloqueos` */

CREATE TABLE `zbloqueos` (
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `tabla` char(20) NOT NULL default '',
  `clave` char(30) NOT NULL default '',
  PRIMARY KEY  (`tabla`,`clave`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `zbloqueos` */

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
