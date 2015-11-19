/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.22 : Database - ariagroutil
*********************************************************************
Server version : 5.0.22
*/ 
SET NAMES utf8;

SET SQL_MODE='';

create database if not exists `ariagroutil`;

USE `ariagroutil`;

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


/*Table structure for table `appmenususuario` */

CREATE TABLE `appmenususuario` (
  `aplicacion` varchar(15) NOT NULL default '0',
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `codigo` smallint(3) unsigned NOT NULL default '0',
  `tag` varchar(100) default '0',
  PRIMARY KEY  (`aplicacion`,`codusu`,`codigo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

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



/*Table structure for table `pcs` */

CREATE TABLE `pcs` (
  `codpc` smallint(5) unsigned NOT NULL default '0',
  `nompc` char(30) default NULL,
  `Conta` smallint(5) unsigned default NULL,
  PRIMARY KEY  (`codpc`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;



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


/*Table structure for table `seglinea` */

CREATE TABLE `seglinea` (
  `codlinea` smallint(3) unsigned NOT NULL,
  `nomlinea` varchar(25) NOT NULL,
  PRIMARY KEY  (`codlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;



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
  `impampli` decimal(8,2) default NULL,
  `impreduc` decimal(8,2) default NULL,
  `intconta` tinyint(1) NOT NULL default '0',
  `inttesor` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`codrefer`,`codiplan`,`codlinea`),
  KEY `FK_segpoliza` (`codlinea`),
  CONSTRAINT `segpoliza_ibfk_1` FOREIGN KEY (`codlinea`) REFERENCES `seglinea` (`codlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;


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

insert into `sempre` (`codempre`,`nomempre`,`domempre`,`codposta`,`pobempre`,`proempre`,`nifempre`,`telempre`,`faxempre`,`wwwempre`,`maiempre`,`perempre`) values (1,'ALZICOOP COOP.V. ','Plaza Mayor, 30','46600','ALZIRA','VALENCIA','F46024196','962404136','962400203','http://www.alzicoop.com','mgarcia@alzicoop.com','Enrique Llorca Martinez');

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
  PRIMARY KEY  (`codparam`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sparam` */

insert into `sparam` (`codparam`,`serconta`,`usuconta`,`pasconta`,`numconta`,`ctagasto`,`ctareten`,`websoporte`,`diremail`,`smtphost`,`smtpuser`,`smtppass`,`porcinte`,`porcrete`,`concedebe`,`concehaber`,`numdiario`,`seguros`,`avnics`,`telefonia`,`sercontaseg`,`usucontaseg`,`pascontaseg`,`numcontaseg`,`concedebeseg`,`concehaberseg`,`numdiarioseg`,`ctabancoseg`,`raizctasocseg`) values (1,'192.168.4.4','root','aritel',7,'430000001','629200001','www.ariadnasoftware.com','monica@myariadna.com','ariadna.myariadna.com','monica','monica98','123.00','15.00',1,2,1,1,1,0,'192.168.4.4','root','aritel',1,1,2,1,'1210095','4300');

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
