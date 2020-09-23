/*
SQLyog Enterprise v4.06 RC1
Host - 5.0.21-community-nt : Database - pa_online
*********************************************************************
Server version : 5.0.21-community-nt
*/


create database if not exists `pa_online`;

USE `pa_online`;

/*Table structure for table `_pa_botredirect` */

drop table if exists `_pa_botredirect`;

CREATE TABLE `_pa_botredirect` (
  `bottagid` int(10) unsigned NOT NULL auto_increment,
  `bottype` enum('msnbot.msn.com','archive.org','googlebot.com','inktomisearch.com','ask.com','looksmart.com') default NULL,
  `actiontype` enum('xhtml','link') default 'link',
  `link` mediumtext,
  `linktext` varchar(128) default NULL,
  `xHTML` text,
  PRIMARY KEY  (`bottagid`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_addresses` */

drop table if exists `acci_addresses`;

CREATE TABLE `acci_addresses` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ContactName` varchar(50) default NULL,
  `Street1` varchar(150) default NULL,
  `Street2` varchar(150) default NULL,
  `Country` varchar(50) default NULL,
  `State` varchar(50) default NULL,
  `Postcode` varchar(20) default NULL,
  `Suburb` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default '-1',
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `AccI_RecID` (`AccI_RecID`),
  KEY `AccountInfoacci_Addresses` (`AccI_RecID`),
  KEY `FlagID` (`FlagID`),
  KEY `Postcode` (`Postcode`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_aliases` */

drop table if exists `acci_aliases`;

CREATE TABLE `acci_aliases` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `acci_RecID` bigint(20) default '0',
  `type` enum('site','system') default 'site',
  `email` varchar(255) default NULL,
  `dest` varchar(255) default NULL,
  `Checked` tinyint(4) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_dslconnections` */

drop table if exists `acci_dslconnections`;

CREATE TABLE `acci_dslconnections` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccountName` varchar(255) default NULL,
  `AreaCode` char(3) default NULL,
  `PhoneNumber` varchar(20) default NULL,
  `eMail` varchar(255) default NULL,
  `acci_RecID` bigint(50) default NULL,
  `Checked` tinyint(4) default NULL,
  `UnitNo` varchar(10) default NULL,
  `StreetNo` varchar(10) default NULL,
  `StreetName` varchar(50) default NULL,
  `StreetType` varchar(20) default NULL,
  `Suburb` varchar(50) default NULL,
  `Country` varchar(50) default NULL,
  `Postcode` varchar(4) default NULL,
  `State` char(3) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `Churn` tinyint(4) default '0',
  `RadiusID` bigint(20) default '0',
  `AccountActive` int(11) default '0',
  `Cflag` tinyint(4) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_editlog` */

drop table if exists `acci_editlog`;

CREATE TABLE `acci_editlog` (
  `RecID` bigint(50) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default NULL,
  `DateEditMade` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SysopID` smallint(6) default NULL,
  `EditTxt` varchar(255) default NULL,
  `IPAddress` varchar(13) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_emailaddresses` */

drop table if exists `acci_emailaddresses`;

CREATE TABLE `acci_emailaddresses` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `EmailAddress` varchar(255) default NULL,
  `ContactName` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default '-1',
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `AccI_RecID` (`AccI_RecID`),
  KEY `AccountInfoacci_EmailAddresses` (`AccI_RecID`),
  KEY `FlagID` (`FlagID`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_flags` */

drop table if exists `acci_flags`;

CREATE TABLE `acci_flags` (
  `RecID` int(4) NOT NULL auto_increment,
  `FlagDesc` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Flag` int(11) default '-1',
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_hardware` */

drop table if exists `acci_hardware`;

CREATE TABLE `acci_hardware` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Modem` varchar(255) default NULL,
  `Processor` varchar(255) default NULL,
  `VideoCard` varchar(255) default NULL,
  `Monitor` varchar(255) default NULL,
  `PCType` varchar(255) default NULL,
  `NetworkCard` varchar(255) default NULL,
  `OS` varchar(255) default NULL,
  `Printer` varchar(255) default NULL,
  `Mainboard` varchar(255) default NULL,
  `acci_RecID` bigint(20) default NULL,
  `acciServiceID` bigint(20) default NULL,
  `ExtraXML` text,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_paymentsettings` */

drop table if exists `acci_paymentsettings`;

CREATE TABLE `acci_paymentsettings` (
  `RecID` int(8) NOT NULL auto_increment,
  `ddAccountName` varchar(255) default NULL,
  `ddBSB` varchar(6) default NULL,
  `ddAccountNumber` varchar(128) default NULL,
  `ddPercentile` float(11,10) default '10.0000000000',
  `ccCardName` varchar(255) default NULL,
  `ccCardNumber` varchar(255) default NULL,
  `ccExpiryDate` varchar(10) default NULL,
  `ccCIC` varchar(10) default NULL,
  `ccType` int(4) default NULL,
  `swWord` varchar(255) default NULL,
  `swNumber` varchar(128) default NULL,
  `cOrder` int(8) default '0',
  `acci_RecID` int(11) default NULL,
  `Checked` int(4) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_phonenumbers` */

drop table if exists `acci_phonenumbers`;

CREATE TABLE `acci_phonenumbers` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `PhoneNumber` varchar(50) default NULL,
  `Extension` varchar(10) default NULL,
  `ContactName` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default '-1',
  `ShortNote` varchar(50) default NULL,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `AccI_RecID` (`AccI_RecID`),
  KEY `AccountInfoacci_PhoneNumbers` (`AccI_RecID`),
  KEY `FlagID` (`FlagID`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_quotareceipt` */

drop table if exists `acci_quotareceipt`;

CREATE TABLE `acci_quotareceipt` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `acci_RecID` bigint(20) default NULL,
  `Description` varchar(255) default NULL,
  `UnitsLeft` int(11) default NULL,
  `UnitsUsed` int(11) default NULL,
  `Flag` smallint(6) default NULL,
  `UnitType` varchar(10) default NULL,
  `QuotaMSGSent` smallint(6) default NULL,
  `QuotaMSGID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_referedby` */

drop table if exists `acci_referedby`;

CREATE TABLE `acci_referedby` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `acci_RecID2` bigint(20) default '0',
  `FlagID` bigint(20) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ContactName` varchar(50) default NULL,
  `ShortNote` varchar(255) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default NULL,
  `acciServiceID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `AccI_RecID` (`AccI_RecID`),
  KEY `Acci_Refered_RecID` (`acci_RecID2`),
  KEY `AccountInfoacci_referedby` (`AccI_RecID`),
  KEY `FlagID` (`FlagID`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `acci_services` */

drop table if exists `acci_services`;

CREATE TABLE `acci_services` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default NULL,
  `ServiceID` bigint(20) default NULL,
  `ContactName` varchar(100) default NULL,
  `Username` varchar(50) default NULL,
  `Password` varchar(50) default NULL,
  `NextCycle` datetime default '0000-00-00 00:00:00',
  `BaseURL` varchar(255) default NULL,
  `RadiusID` bigint(20) default NULL,
  `DateCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `AccI_RecID` bigint(20) default NULL,
  `DynamicField1` varchar(255) default NULL,
  `DynamicField2` varchar(255) default NULL,
  `DynamicField3` varchar(255) default NULL,
  `DynamicField4` varchar(255) default NULL,
  `DynamicField5` varchar(255) default NULL,
  `Checked` tinyint(4) default '0',
  `VirtualID` bigint(20) default NULL,
  `PreviousCycle` datetime default '0000-00-00 00:00:00',
  `PreviousCycleA` datetime default '0000-00-00 00:00:00',
  `PreviousCycleB` datetime default '0000-00-00 00:00:00',
  `PreviousCycleC` datetime default '0000-00-00 00:00:00',
  `PreviousCycleD` datetime default '0000-00-00 00:00:00',
  `PreviousCycleE` datetime default '0000-00-00 00:00:00',
  `PreviousCycleF` datetime default '0000-00-00 00:00:00',
  `PreviousCycleG` datetime default '0000-00-00 00:00:00',
  `DomainID` bigint(20) default '0',
  `SysopID` bigint(20) default '1',
  `SubRecID` bigint(20) default '0',
  `MBQuota` int(11) default '10',
  `Activation` datetime default NULL,
  `ActivationSet` int(8) default '0',
  `SystemUID` smallint(5) unsigned default '0',
  `PeriodFee` float(13,5) default '0.00000',
  `PerHour` float(13,5) default '0.00000',
  `PerMB` float(13,5) default '0.00000',
  `JoiningFee` float(13,5) default '0.00000',
  `AgencyID` bigint(20) default '0',
  `POID` bigint(20) default '0',
  `DefaultShippingID` bigint(20) default '0',
  `ContractID` int(11) default NULL,
  `ContractExpiry` datetime default NULL,
  `crsLocal` float(31,30) default '0.000000000000000000000000000000',
  `crsNational` float(31,30) default '0.000000000000000000000000000000',
  `crsInternational` float(31,30) default '0.000000000000000000000000000000',
  `crsMobile` float(31,30) default '0.000000000000000000000000000000',
  `crsVOIP` float(31,30) default '0.000000000000000000000000000000',
  `crsSpecial` float(31,30) default '0.000000000000000000000000000000',
  `crsFlatRate` float(31,30) default '0.000000000000000000000000000000',
  `crsRateMeasured` float(31,30) default '0.000000000000000000000000000000',
  `crsCallCredits` float(31,30) default '0.000000000000000000000000000000',
  `creLocal` float(31,30) default '0.000000000000000000000000000000',
  `creNational` float(31,30) default '0.000000000000000000000000000000',
  `creInternational` float(31,30) default '0.000000000000000000000000000000',
  `creMobile` float(31,30) default '0.000000000000000000000000000000',
  `creVOIP` float(31,30) default '0.000000000000000000000000000000',
  `creSpecial` float(31,30) default '0.000000000000000000000000000000',
  `creFlatRate` float(31,30) default '0.000000000000000000000000000000',
  `creRateMeasured` char(2) default 'n',
  `creRateMeasuredphp` varchar(26) default '1 minute',
  `CallCredits` float(31,30) default '0.000000000000000000000000000000',
  `LineRental` float(31,30) default '0.000000000000000000000000000000',
  `EquipmentHire` float(31,30) default '0.000000000000000000000000000000',
  `ExtendedServiceCost` float(31,30) default '0.000000000000000000000000000000',
  `CallCapping` float(31,30) default '0.000000000000000000000000000000',
  `FlagFall` float(31,30) default '0.000000000000000000000000000000',
  `Quanity` int(11) default '1',
  `PHPSESSION` varchar(255) default 'NULL',
  `SysopIDRunningMaitenance` bigint(20) default '0',
  `SysopNetworkHostname` tinytext,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `accountclass` */

drop table if exists `accountclass`;

CREATE TABLE `accountclass` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Description` varchar(50) NOT NULL default '',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `accountinfo` */

drop table if exists `accountinfo`;

CREATE TABLE `accountinfo` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccountName` varchar(255) default NULL,
  `ActivationDate` datetime default NULL,
  `ExpiryDate` datetime default NULL,
  `CreationDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `DOB` datetime default NULL,
  `PaymentID` bigint(20) default NULL,
  `ActivateDateSet` smallint(6) default NULL,
  `ExpiryDateSet` smallint(6) default NULL,
  `ProcessFlag` smallint(6) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default '-1',
  `Classification` int(11) default '1',
  `Realm` varchar(255) default NULL,
  `PayIntervalType` char(2) default 'd',
  `PayInterval` int(11) default '14',
  `FlagA_RecID` bigint(20) default '1',
  `FlagB_RecID` bigint(20) default '7',
  `SysopID` bigint(20) default '1',
  `VirtualID` bigint(20) default NULL,
  `AgencyID` bigint(20) default NULL,
  `sfStartTime` datetime default NULL,
  `sfCycle_Upload` bigint(20) default '0',
  `sfCycle_Download` bigint(20) default '0',
  `sfCycle_Mins` bigint(20) default '0',
  `Crystallise` tinyint(4) default '0',
  `Omnivorous` tinyint(4) default '0',
  `Carnivorous` tinyint(4) default '0',
  `FlagASet` datetime default NULL,
  `FLagBSet` datetime default NULL,
  `BillingDate` datetime default NULL,
  `gUsername` varchar(50) default NULL,
  `gPassword` varchar(50) default NULL,
  `AboutUs` tinyint(4) default '0',
  `ftpPathKey` varchar(128) default NULL,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  UNIQUE KEY `gUsernameUNIQUE` (`gUsername`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `accountviewer` */

drop table if exists `accountviewer`;

CREATE TABLE `accountviewer` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `SubofRecID` bigint(20) default NULL,
  `Description` varchar(50) default NULL,
  `IconNum` int(11) default '42',
  `Action` varchar(50) default '0',
  `SelectStatement` text,
  `CountStatement` text,
  `VirtualID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `agency` */

drop table if exists `agency`;

CREATE TABLE `agency` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AgencyName` varchar(255) default NULL,
  `CreatedBy` bigint(20) default NULL,
  `SysopID` bigint(20) default NULL,
  `Icon` int(11) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `Contract1` tinyint(4) default '0',
  `Contract2` tinyint(4) default '0',
  `Contract3` tinyint(4) default '0',
  `Contract4` tinyint(4) default '0',
  `Comment` text,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `agencyplans` */

drop table if exists `agencyplans`;

CREATE TABLE `agencyplans` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `sKey` char(50) default NULL,
  `ptRecID` bigint(20) default NULL,
  `IsAvailable` tinyint(4) default NULL,
  `AgencyID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `audithistory` */

drop table if exists `audithistory`;

CREATE TABLE `audithistory` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `systemstamp` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `sysnow` datetime default NULL,
  `localtime` datetime default NULL,
  `appname` varchar(128) default NULL,
  `appversion` varchar(32) default '2.9.0.330',
  `apphdc` int(8) default '0',
  `formname` varchar(128) default NULL,
  `formhwnd` int(8) default '0',
  `vartype` varchar(32) NOT NULL default '',
  `varname` varchar(64) default NULL,
  `oldcur` float(31,6) default '0.000000',
  `newcur` float(31,6) default '0.000000',
  `oldvalue` text,
  `newvalue` text,
  `oldpointer` bigint(20) default '0',
  `newpointer` bigint(20) default '0',
  `IDX` int(8) default '0',
  `sysopid` bigint(20) default '0',
  `virtualid` bigint(20) default '0',
  `agencyid` bigint(20) default '0',
  `acci_RecID` bigint(20) default '0',
  `refundid` bigint(20) default '0',
  `invtrxrid` bigint(20) default '0',
  `flagid` int(4) default '0',
  `Description` varchar(255) default NULL,
  `Checked` tinyint(2) default '-1',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

/*Table structure for table `automessages` */

drop table if exists `automessages`;

CREATE TABLE `automessages` (
  `RecID` int(11) NOT NULL auto_increment,
  `MSGType` int(11) default NULL,
  `MessageDraft` text,
  `Description` varchar(50) default NULL,
  `HTMLDraft` text,
  `VirtualID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `bmessages` */

drop table if exists `bmessages`;

CREATE TABLE `bmessages` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `acci_RecID` bigint(20) default NULL,
  `Data90` tinyint(4) default NULL,
  `Time90` tinyint(50) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `bonus_awards` */

drop table if exists `bonus_awards`;

CREATE TABLE `bonus_awards` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default '0',
  `AwardTo` int(11) default '0',
  `Min` int(11) default '0',
  `Max` int(11) default '0',
  `UnitType` int(11) default '0',
  `Units` int(11) default '0',
  `VirtualID` bigint(20) default '0',
  `sKey` char(50) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `bonus_matrix` */

drop table if exists `bonus_matrix`;

CREATE TABLE `bonus_matrix` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AwardTo` int(11) default '0',
  `TransactionText` char(255) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SysopID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `AgencyID` bigint(20) default '0',
  `DeveloperID` bigint(20) default '0',
  `Credit` bigint(20) default '0',
  `DigitalCredit` bigint(20) default '0',
  `acci_RecID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `bonus_units` */

drop table if exists `bonus_units`;

CREATE TABLE `bonus_units` (
  `RecID` int(11) NOT NULL default '0',
  `UnitName` char(50) default NULL,
  `FieldName` char(30) default NULL,
  `Format` char(64) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `cc_receipt` */

drop table if exists `cc_receipt`;

CREATE TABLE `cc_receipt` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `cc_RecID` bigint(20) default NULL,
  `ReceiptNumber` varchar(50) default NULL,
  `RecCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `cctransactions` */

drop table if exists `cctransactions`;

CREATE TABLE `cctransactions` (
  `RecID` bigint(20) NOT NULL default '0',
  `Amount` float(13,5) default NULL,
  `CCRecID` bigint(20) default NULL,
  `CurrencyCode` varchar(10) default NULL,
  `ECIType` int(11) default NULL,
  `ErrorCode` int(11) default NULL,
  `ErrorMessage` varchar(255) default NULL,
  `ID` int(11) default NULL,
  `OwnerID` int(11) default NULL,
  `ProcessDate` datetime default NULL,
  `ReceiptNumber` varchar(50) default NULL,
  `RequesterIPAddress` varchar(15) default NULL,
  `ResponseText` varchar(255) default NULL,
  `SecureCode` varchar(50) default NULL,
  `Status` int(11) default NULL,
  `SummaryCode` int(11) default NULL,
  `TransactionCode` varchar(50) default NULL,
  `TransactionType` int(11) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `columnlayout` */

drop table if exists `columnlayout`;

CREATE TABLE `columnlayout` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Description` varchar(50) default NULL,
  `ServiceKey` varchar(50) default NULL,
  `Width` int(11) default '1440',
  `FieldName` varchar(100) default NULL,
  `sFormat` tinytext,
  `cOrder` int(11) default '0',
  `formcode` varchar(64) default NULL,
  `SECLevel` int(11) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `commnames` */

drop table if exists `commnames`;

CREATE TABLE `commnames` (
  `RecID` int(11) NOT NULL auto_increment,
  `CommName` varchar(50) default NULL,
  `VirtualID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `commrates` */

drop table if exists `commrates`;

CREATE TABLE `commrates` (
  `RecID` int(11) NOT NULL auto_increment,
  `ClassID` int(11) default NULL,
  `ptRecID` bigint(20) default NULL,
  `Min` int(11) default NULL,
  `Max` int(11) default NULL,
  `CommPerUnit` float(13,5) default NULL,
  `Margin` float(13,5) default NULL,
  `Percentile` float(13,5) default NULL,
  `VirtualID` bigint(20) default NULL,
  `sKey` varchar(50) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `contractsruntime` */

drop table if exists `contractsruntime`;

CREATE TABLE `contractsruntime` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default NULL,
  `TemplateID` bigint(20) default NULL,
  `ContractID` int(11) default NULL,
  `VirtualID` bigint(20) default NULL,
  `Termination` float(31,30) default NULL,
  `JoiningFee` float(31,30) default NULL,
  `PeriodFee` float(31,30) default NULL,
  `FeePerBlock` float(31,30) default NULL,
  `FeePerHour` float(31,30) default NULL,
  `SysopID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `contracttemplates` */

drop table if exists `contracttemplates`;

CREATE TABLE `contracttemplates` (
  `RecID` int(11) NOT NULL auto_increment,
  `ptRecID` bigint(20) default '0',
  `Description` char(128) default '(null)',
  `NoPeriods` int(11) default '52',
  `TypePeriods` char(10) default 'w',
  `Termination` float(31,30) default '0.000000000000000000000000000000',
  `PeriodFee` float(31,30) default '0.000000000000000000000000000000',
  `FeePerBlock` float(31,30) default '0.000000000000000000000000000000',
  `FeePerHour` float(31,30) default '0.000000000000000000000000000000',
  `JoiningFee` float(31,30) default '0.000000000000000000000000000000',
  `bDeleted` tinyint(4) default '0',
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `creditcard` */

drop table if exists `creditcard`;

CREATE TABLE `creditcard` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `CardNumber` varchar(255) default '0',
  `SecurityNumber` varchar(255) default '0',
  `ExpiryDate` date default '0000-00-00',
  `Name` varchar(50) default '"Not Set"',
  `bDefault` tinyint(4) default '-1',
  `bType` tinyint(4) default '0',
  `VirtualID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `dnszones` */

drop table if exists `dnszones`;

CREATE TABLE `dnszones` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Class` varchar(5) default 'IN',
  `Type` varchar(6) default 'CNAMES',
  `Data` text,
  `SerialNumber` bigint(20) default NULL,
  `Checked` int(11) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `domaindocs` */

drop table if exists `domaindocs`;

CREATE TABLE `domaindocs` (
  `RecID` bigint(20) unsigned NOT NULL auto_increment,
  `DomainID` bigint(20) default NULL,
  `DocType` varchar(64) default NULL,
  `DocText` text,
  `Icon` tinyint(4) default NULL,
  `Description` varchar(255) default NULL,
  `ItemText` varchar(64) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `domainlist` */

drop table if exists `domainlist`;

CREATE TABLE `domainlist` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Domain` varchar(255) default NULL,
  `AdminEmail` varchar(255) default NULL,
  `Status` char(1) default NULL,
  `Checked` int(11) default NULL,
  `SubDomainLimit` int(11) default '1',
  `vKey` varchar(255) default NULL,
  `acci_RecID` bigint(20) default '0',
  `TechName` varchar(255) default NULL,
  `TechPass` varchar(128) default NULL,
  `SysopID` bigint(20) default NULL,
  `VirtualID` bigint(20) default NULL,
  `ContactName` varchar(128) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `errorlog` */

drop table if exists `errorlog`;

CREATE TABLE `errorlog` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ErrNo` bigint(20) default NULL,
  `sDesc` text,
  `LBL` text,
  `iAction` int(12) default NULL,
  `Container` varchar(180) default NULL,
  `Routine` varchar(180) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `LastOCC` datetime default NULL,
  `Username` varchar(128) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `exp_categories` */

drop table if exists `exp_categories`;

CREATE TABLE `exp_categories` (
  `RecID` bigint(20) NOT NULL default '0',
  `SubRecID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `Icon` int(4) default '1',
  `Description` char(255) default 'Not Set (null)',
  `formcode` char(64) default 'exp001',
  `SecLevel` tinyint(4) default '50',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `exp_flags` */

drop table if exists `exp_flags`;

CREATE TABLE `exp_flags` (
  `RecID` bigint(64) NOT NULL default '0',
  `InBookID` bigint(32) default '0',
  `Flag` int(11) default '0',
  `AmountPaid` float(31,30) default '0.000000000000000000000000000000',
  `RemittanceSent` tinyint(4) default '0',
  `PaymentMethod` bigint(20) default '0',
  `SerialNo` char(64) default NULL,
  `PaidTo` char(255) default NULL,
  `TransactionText` char(255) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `exp_inbook` */

drop table if exists `exp_inbook`;

CREATE TABLE `exp_inbook` (
  `RecID` bigint(32) NOT NULL default '0',
  `VirtualID` bigint(20) default NULL,
  `SysopID` bigint(20) default NULL,
  `CategoryID` bigint(20) default NULL,
  `Description` char(255) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `AssVirtualID` bigint(20) default NULL,
  `AssSysopID` bigint(20) default NULL,
  `AmountDUE` float(31,30) default NULL,
  `GST` float(31,30) default NULL,
  `AmountPaid` float(31,30) default NULL,
  `dBodyName` char(255) default NULL,
  `dContactname` char(255) default NULL,
  `dPhoneNumber` char(255) default NULL,
  `dFaxNumber` char(255) default NULL,
  `dAddress` char(255) default NULL,
  `dEmailAddress` char(255) default NULL,
  `dAccNo` char(64) default NULL,
  `dInvoiceNo` char(64) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags` */

drop table if exists `flags`;

CREATE TABLE `flags` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `FlagDesc` varchar(50) default NULL,
  `rolloverIntervalType` char(1) default NULL,
  `rolloverInterval` int(11) default NULL,
  `rollover2FlagID` bigint(20) default NULL,
  `ListedOnRadius` smallint(6) default '-1',
  `FlagType` char(1) default 'a',
  `FlagCode` varchar(30) default NULL,
  `Icon` varchar(64) default NULL,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`),
  KEY `RecID` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_invoicein` */

drop table if exists `flags_invoicein`;

CREATE TABLE `flags_invoicein` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvIn_RecID` bigint(20) default NULL,
  `Flag` bigint(20) default NULL,
  `SysopID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_invoiceout` */

drop table if exists `flags_invoiceout`;

CREATE TABLE `flags_invoiceout` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvOut_RecID` bigint(20) default NULL,
  `Flag` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_invoicetrx` */

drop table if exists `flags_invoicetrx`;

CREATE TABLE `flags_invoicetrx` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `FlagID` bigint(20) default NULL,
  `InvTrx_RecID` bigint(50) default NULL,
  `SysopID` bigint(20) default NULL,
  `VirtualID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_plantype` */

drop table if exists `flags_plantype`;

CREATE TABLE `flags_plantype` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default NULL,
  `NumberOf` tinyint(4) default NULL,
  `Checked` tinyint(4) default NULL,
  `PlanType` bigint(20) default NULL,
  `VirtualID` bigint(20) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_refunds` */

drop table if exists `flags_refunds`;

CREATE TABLE `flags_refunds` (
  `RecID` bigint(20) NOT NULL default '0',
  `acciServiceID` bigint(20) default '0',
  `Refunded` float(31,30) default '0.000000000000000000000000000000',
  `GST` float(31,30) default '0.000000000000000000000000000000',
  `SysopID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `acci_RecID` bigint(20) default '0',
  `RefundTRXID` bigint(20) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `flags_tempextras` */

drop table if exists `flags_tempextras`;

CREATE TABLE `flags_tempextras` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default NULL,
  `NumberOf` tinyint(4) default NULL,
  `Checked` tinyint(4) default NULL,
  `PlanType` bigint(20) default NULL,
  `ContractID` int(11) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `history_acci_datausage` */

drop table if exists `history_acci_datausage`;

CREATE TABLE `history_acci_datausage` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `acci_RecID` bigint(20) default '0',
  `Uploaded` mediumint(9) default '0',
  `Downloaded` mediumint(9) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `NumMins` mediumint(9) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Table stores history of Data Usage to build datawave graph';

/*Table structure for table `history_radius_datausage` */

drop table if exists `history_radius_datausage`;

CREATE TABLE `history_radius_datausage` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `RadiusID` bigint(20) default '0',
  `Uploaded` mediumint(9) default '0',
  `Downloaded` mediumint(9) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `NumMins` mediumint(9) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Table stores history of Data Usage to build datawave graph';

/*Table structure for table `imappop` */

drop table if exists `imappop`;

CREATE TABLE `imappop` (
  `id` varchar(60) default NULL,
  `clear` varchar(30) NOT NULL default '',
  `crypt` varchar(255) NOT NULL default '',
  `uid` int(11) default NULL,
  `gid` int(11) default NULL,
  `name` varchar(255) default NULL,
  `home` varchar(255) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `invoicein` */

drop table if exists `invoicein`;

CREATE TABLE `invoicein` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default NULL,
  `AmountPaid` float(12,2) default '0.00',
  `PaidWhen` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `Checked` int(11) default '0',
  `AmountUsed` float(12,2) default '0.00',
  `GSTCharged` float(12,2) default '0.00',
  `Sub` float(12,2) unsigned zerofill default '000000000.00',
  `TotalPaid` float(12,2) unsigned zerofill default '000000000.00',
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `RefundTRXID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `invoiceout` */

drop table if exists `invoiceout`;

CREATE TABLE `invoiceout` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AccI_RecID` bigint(20) default '0',
  `AmountDue` float(11,2) default '0.00',
  `GSTCharged` float(11,2) default '0.00',
  `PaymentDue` datetime default NULL,
  `AmountPaid` float(11,2) default '0.00',
  `PaidWhen` datetime default NULL,
  `Checked` int(11) default '0',
  `FlagID` int(11) default '0',
  `TotalDue` float(11,2) default '0.00',
  `sfCycle_Upload` bigint(20) default '0',
  `sfCycle_Download` bigint(20) default '0',
  `sfCycle_Mins` bigint(20) default '0',
  `AgencyID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `Description` varchar(255) default NULL,
  `TraxrID` bigint(20) default '0',
  `RefundID` bigint(20) default '0',
  `PlanServiceID` bigint(50) default '0',
  `AmountRefunded` float(11,2) default '0.00',
  `GSTRefunded` float(11,2) default '0.00',
  `SysopID` bigint(20) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SubRecID` bigint(20) default '0',
  `VISPStatementID` bigint(20) default '0',
  `StatementID` bigint(20) default '0',
  `InvoiceFlagID` varchar(20) default '0',
  `StartCycle` datetime default '0000-00-00 00:00:00',
  `EndCycle` datetime default '0000-00-00 00:00:00',
  `ptRecID` bigint(20) default '0',
  `ServiceID` bigint(20) default '0',
  `DomainID` bigint(20) default '0',
  `RadiusID` bigint(20) default '0',
  `Quanity` int(11) default '1',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `invoicetraxr` */

drop table if exists `invoicetraxr`;

CREATE TABLE `invoicetraxr` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvoiceSerial` varchar(50) default '0',
  `acci_RecID` bigint(20) default NULL,
  `TotalDue` float(42,2) default '0.00',
  `AmountPaid` float(42,2) default '0.00',
  `PaymentDue` datetime default NULL,
  `Finalised` tinyint(4) default '0',
  `Comment` text,
  `PaidWhen` datetime default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `InvoiceID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `AgencyID` bigint(20) default '0',
  `StatementID` double(13,5) default NULL,
  `AmountCredited` float(42,2) default '0.00',
  `GSTDue` float(31,2) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `invout_payment` */

drop table if exists `invout_payment`;

CREATE TABLE `invout_payment` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvOut_RecID` bigint(20) default '0',
  `Amount` float(12,2) default '0.00',
  `GST` float(12,2) default '0.00',
  `Sub` float(12,2) default '0.00',
  `WhenPaid` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `AccI_RecID` bigint(20) default NULL,
  `TotalPaid` float(12,2) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `invtrx_payment` */

drop table if exists `invtrx_payment`;

CREATE TABLE `invtrx_payment` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `WhenPaid` datetime default NULL,
  `Amount` float(5,2) default NULL,
  `GST` float(5,2) default NULL,
  `Sub` float(5,2) default NULL,
  `InvTrxRecID` bigint(20) default NULL,
  `acci_RecID` bigint(20) default NULL,
  `SysopID` bigint(20) default NULL,
  `VirtualID` bigint(20) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `irccertificates` */

drop table if exists `irccertificates`;

CREATE TABLE `irccertificates` (
  `RecID` int(11) NOT NULL auto_increment,
  `A` text,
  `B` text,
  `nick` varchar(64) default NULL,
  `systime` datetime default NULL,
  `localtime` datetime default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `ircservices` */

drop table if exists `ircservices`;

CREATE TABLE `ircservices` (
  `ID` int(8) NOT NULL auto_increment,
  `server` varchar(255) default NULL,
  `port` int(8) default '6667',
  `channel` varchar(64) default '#projectalpha',
  `countrycode` varchar(16) default 'AUS',
  PRIMARY KEY  (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `news_plantemplates` */

drop table if exists `news_plantemplates`;

CREATE TABLE `news_plantemplates` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ptRecID` bigint(20) default '0',
  `NewsID` bigint(20) default '0',
  `ServiceID` bigint(20) default '0',
  `VendorID` int(11) default '0',
  `bSharingLevel` tinyint(4) default '1',
  `PeriodFee` float(31,30) default '0.000000000000000000000000000000',
  `MBPerPeriod` int(11) NOT NULL default '0',
  `MBBlockSize` int(11) default '0',
  `FeePerBlock` float(31,30) default '0.000000000000000000000000000000',
  `HoursPerPeriod` int(11) default '0',
  `ExtraPerHour` float(31,30) default '0.000000000000000000000000000000',
  `MBQuota` int(11) default '10',
  `VendorPartID` varchar(30) default NULL,
  `SubPartID` varchar(30) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `VirtualID` bigint(20) default '0',
  `CategoryID` bigint(20) default '0',
  `crsLocal` float(31,30) default '0.500000000000000000000000000000',
  `crsNational` float(31,30) default '0.500000000000000000000000000000',
  `crsInternational` float(31,30) default '0.500000000000000000000000000000',
  `crsMobile` float(31,30) default '0.500000000000000000000000000000',
  `crsVOIP` float(31,30) default '0.500000000000000000000000000000',
  `crsSpecial` float(31,30) default '0.500000000000000000000000000000',
  `crsRateMeasure` char(2) default 'n',
  `creRateMeasuredphp` varchar(26) default '1 minute',
  `crsFlatRate` tinyint(4) default '0',
  `crsCallCredits` float(31,30) default '0.000000000000000000000000000000',
  `creLocal` float(31,30) default '0.500000000000000000000000000000',
  `creNational` float(31,30) default '0.500000000000000000000000000000',
  `creInternational` float(31,30) default '0.500000000000000000000000000000',
  `creMobile` float(31,30) default '0.500000000000000000000000000000',
  `creVOIP` float(31,30) default '0.500000000000000000000000000000',
  `creSpecial` float(31,30) default '0.500000000000000000000000000000',
  `creRateMeasure` char(2) default 'm',
  `creFlatRate` tinyint(4) default '0',
  `CallCredits` float(31,30) default '0.000000000000000000000000000000',
  `LineRental` float(31,30) default '0.000000000000000000000000000000',
  `EquipmentHire` float(31,30) default '0.000000000000000000000000000000',
  `ExtendedServiceCost` float(31,30) default '0.000000000000000000000000000000',
  `CallCapping` float(31,30) default '0.000000000000000000000000000000',
  `FlagFall` float(31,30) default '0.000000000000000000000000000000',
  `OrderQuanity` int(11) default '1',
  `Stock` int(11) default '0',
  `WarehouseQuanityA` int(11) default '0',
  `PickingQuanityA` int(11) default '0',
  `FetchCode` varchar(42) default NULL,
  PRIMARY KEY  (`RecID`),
  KEY `ServiceID` (`ServiceID`,`VendorID`,`VendorPartID`,`SubPartID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `news_plantypes` */

drop table if exists `news_plantypes`;

CREATE TABLE `news_plantypes` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ServiceID` bigint(20) default '0',
  `VendorID` int(11) default '0',
  `CatNo` char(30) default 'I0000',
  `Barcode` char(30) default NULL,
  `Description` char(128) default NULL,
  `PeriodFee` float(31,30) default '0.000000000000000000000000000000',
  `MBPerPeriod` int(11) NOT NULL default '0',
  `MBBlockSize` int(11) default '0',
  `FeePerBlock` float(31,30) default '0.000000000000000000000000000000',
  `HoursPerPeriod` int(11) default '0',
  `ExtraPerHour` float(31,30) default '0.000000000000000000000000000000',
  `Rollover` tinyint(4) default '0',
  `roIntervalType` char(1) default 'm',
  `roInterval` int(11) default '1',
  `roRecID` bigint(20) default NULL,
  `chgIntervalType` char(1) default 'm',
  `chgInterval` int(11) default '1',
  `chgIntervalDays` float(11,8) default '31.00000000',
  `chgPerTxt` char(255) default 'day',
  `chgIntervalphp` char(26) default '1 month',
  `BillImmediately` tinyint(4) default '0',
  `VirtualID` bigint(20) default NULL,
  `RadiusID` bigint(20) default NULL,
  `SessionTimeout` int(11) default '10800',
  `IdleTimeout` int(11) default '600',
  `SessionsAllowed` int(11) default '1',
  `TemplateID` bigint(20) default '0',
  `JoiningFee` float(31,30) default '0.000000000000000000000000000000',
  `CostPrice` float(31,30) default '0.000000000000000000000000000000',
  `MBQuota` int(11) default '10',
  `BillOnce` int(11) default '0',
  `ShowOnWeb` enum('Y','N') NOT NULL default 'N',
  `BonusID` int(11) default '0',
  `OptionalText` char(128) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `crsLocal` float(31,30) default '0.500000000000000000000000000000',
  `crsNational` float(31,30) default '0.500000000000000000000000000000',
  `crsInternational` float(31,30) default '0.500000000000000000000000000000',
  `crsMobile` float(31,30) default '0.500000000000000000000000000000',
  `crsVOIP` float(31,30) default '0.500000000000000000000000000000',
  `crsSpecial` float(31,30) default '0.500000000000000000000000000000',
  `crsFlatRate` tinyint(4) default '0',
  `crsRateMeasured` char(2) default 'n',
  `crsCallCredits` float(31,30) default '0.000000000000000000000000000000',
  `creLocal` float(31,30) default '0.500000000000000000000000000000',
  `creNational` float(31,30) default '0.500000000000000000000000000000',
  `creInternational` float(31,30) default '0.500000000000000000000000000000',
  `creMobile` float(31,30) default '0.500000000000000000000000000000',
  `creVOIP` float(31,30) default '0.500000000000000000000000000000',
  `creSpecial` float(31,30) default '0.500000000000000000000000000000',
  `creFlatRate` tinyint(4) default '0',
  `creRateMeasured` char(2) default 'm',
  `creRateMeasuredphp` char(26) default '1 minute',
  `CallCredits` float(31,30) default '0.000000000000000000000000000000',
  `LineRental` float(31,30) default '0.000000000000000000000000000000',
  `EquipmentHire` float(31,30) default '0.000000000000000000000000000000',
  `ExtendedServiceCost` float(31,30) default '0.000000000000000000000000000000',
  `CallCapping` float(31,30) default '0.000000000000000000000000000000',
  `FlagFall` float(31,30) default '0.000000000000000000000000000000',
  `Stock` int(11) default '0',
  `QuanityStock` int(11) default '0',
  `WarehouseQuanity` int(11) default '0',
  `bSharingLevel` tinyint(4) default '1',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `newsfeed` */

drop table if exists `newsfeed`;

CREATE TABLE `newsfeed` (
  `NewsID` int(11) unsigned NOT NULL auto_increment,
  `VirtualID` bigint(20) default '0',
  `aaHeading` varchar(128) default NULL,
  `aaHeadingText-Size` char(3) default '14',
  `aaHeadingText-Color` varchar(6) default 'FC3333',
  `aaDescription` text,
  `aaLink1URL` text,
  `aaLink1Desc` varchar(128) default '(Click Here to See More Details)',
  `aaLink1Colour` varchar(6) default 'CC0000',
  `aaText-Size` char(3) default '10',
  `aaText-Color` varchar(6) default '006633',
  `aaShift-Pause` varchar(6) default '9000',
  `aaShift-In-Effect` varchar(128) default 'slide-left',
  `aaBackground-Color` varchar(6) default 'ffffff',
  `aaBorder-Color` varchar(6) default '66CCFF',
  `escalation` int(4) default '0',
  `moddate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ExpiryDate` datetime default NULL,
  PRIMARY KEY  (`NewsID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `ntpservers` */

drop table if exists `ntpservers`;

CREATE TABLE `ntpservers` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ntpServer` varchar(255) default NULL,
  `location` varchar(255) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `online_maintenancelocker` */

drop table if exists `online_maintenancelocker`;

CREATE TABLE `online_maintenancelocker` (
  `SvrRecID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `InvRecID` bigint(20) default '0',
  `StatementID` bigint(20) default '0',
  `Active` tinyint(2) default '1',
  `ModifiedDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SysopIDRunningMaintenance` bigint(20) default '0',
  `DateNava` datetime default NULL,
  `DateNavb` datetime default NULL,
  `SysopHostName` tinytext,
  `PHPSESSION` varchar(255) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `online_maintenancelockerb` */

drop table if exists `online_maintenancelockerb`;

CREATE TABLE `online_maintenancelockerb` (
  `SvrRecID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `InvRecID` bigint(20) default '0',
  `StatementID` bigint(20) default '0',
  `Active` tinyint(2) default '1',
  `ModifiedDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SysopIDRunningMaintenance` bigint(20) default '0',
  `DateNava` datetime default NULL,
  `DateNavb` datetime default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `online_settingsandtimers` */

drop table if exists `online_settingsandtimers`;

CREATE TABLE `online_settingsandtimers` (
  `ConfigKey` varchar(32) default NULL,
  `IntegerSet` int(11) default NULL,
  `BigIntSet` int(30) default NULL,
  `VarCharset` varchar(255) default NULL,
  `textSet` text,
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `paymenttype` */

drop table if exists `paymenttype`;

CREATE TABLE `paymenttype` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Description` varchar(50) default NULL,
  `Sub` float(12,2) default '0.00',
  `CreditCard` int(4) default '0',
  `HasBSB` tinyint(4) default '0',
  `HasAcc` tinyint(4) default '0',
  `HasName` tinyint(4) default '0',
  `HasSerial` tinyint(4) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `plantemplates` */

drop table if exists `plantemplates`;

CREATE TABLE `plantemplates` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `ServiceID` bigint(20) default '0',
  `VendorID` int(11) default '0',
  `bSharingLevel` tinyint(4) default '1',
  `BarCode` varchar(75) default NULL,
  `Description` varchar(128) default 'Service Or Plan',
  `PeriodFee` float(31,30) default '0.000000000000000000000000000000',
  `MBPerPeriod` int(11) NOT NULL default '0',
  `MBBlockSize` int(11) default '0',
  `FeePerBlock` float(31,30) default '0.000000000000000000000000000000',
  `HoursPerPeriod` int(11) default '0',
  `ExtraPerHour` float(31,30) default '0.000000000000000000000000000000',
  `BillImmediately` tinyint(4) default '0',
  `RadiusID` bigint(20) default NULL,
  `SessionTimeout` int(11) default '10800',
  `IdleTimeout` int(11) default '600',
  `SessionsAllowed` int(11) default '1',
  `CostPrice` float(31,30) default '0.000000000000000000000000000000',
  `MBCostPrice` float(31,30) default '0.000000000000000000000000000000',
  `PeriodCostPrice` float(31,30) default '0.000000000000000000000000000000',
  `Hidden` tinyint(4) default '0',
  `MBQuota` int(11) default '10',
  `VendorPartID` varchar(30) default NULL,
  `SubPartID` varchar(30) default NULL,
  `ProductText` longtext,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `VirtualID` bigint(20) default '0',
  `CategoryID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `height` float(20,5) default '0.00000',
  `lenght` float(20,5) default '0.00000',
  `weight` float(31,30) default '0.000000000000000000000000000000',
  `depth` float(20,5) default '0.00000',
  `packaging` varchar(255) default NULL,
  `unitperpack` int(11) default '1',
  `cFragile` tinyint(2) default '0',
  `cElectrical` tinyint(2) default '0',
  `cAccessory` tinyint(2) default '0',
  `cBotanical` tinyint(2) default '0',
  `cHazardous` tinyint(2) default '0',
  `cChemical` tinyint(2) default '0',
  `cRoad_freight` tinyint(2) default '0',
  `cAir_freight` tinyint(2) default '0',
  `cToy` tinyint(2) default '0',
  `cComponent` tinyint(2) default '0',
  `cComsumable` tinyint(2) default '0',
  `cFood_Stuffs` tinyint(2) default '0',
  `cPrinted_media` tinyint(2) default '0',
  `cSoftware` tinyint(2) default '0',
  `cMembership` tinyint(2) default '0',
  `cHardware` tinyint(2) default '0',
  `cAdult_material` tinyint(2) default '0',
  `cTools` tinyint(2) default '0',
  `location` varchar(255) default NULL,
  `chargebyrate` tinyint(2) default '0',
  `ratetype` char(2) default NULL,
  `rateinterval` int(4) default '0',
  `ratephp` varchar(26) default '1 month',
  `crsLocal` float(31,30) default '0.500000000000000000000000000000',
  `crsNational` float(31,30) default '0.500000000000000000000000000000',
  `crsInternational` float(31,30) default '0.500000000000000000000000000000',
  `crsMobile` float(31,30) default '0.500000000000000000000000000000',
  `crsVOIP` float(31,30) default '0.500000000000000000000000000000',
  `crsSpecial` float(31,30) default '0.500000000000000000000000000000',
  `crsRateMeasure` char(2) default 'n',
  `creRateMeasuredphp` varchar(26) default '1 minute',
  `crsFlatRate` tinyint(4) default '0',
  `crsCallCredits` float(31,30) default '0.000000000000000000000000000000',
  `creLocal` float(31,30) default '0.500000000000000000000000000000',
  `creNational` float(31,30) default '0.500000000000000000000000000000',
  `creInternational` float(31,30) default '0.500000000000000000000000000000',
  `creMobile` float(31,30) default '0.500000000000000000000000000000',
  `creVOIP` float(31,30) default '0.500000000000000000000000000000',
  `creSpecial` float(31,30) default '0.500000000000000000000000000000',
  `creRateMeasure` char(2) default 'm',
  `creFlatRate` tinyint(4) default '0',
  `CallCredits` float(31,30) default '0.000000000000000000000000000000',
  `LineRental` float(31,30) default '0.000000000000000000000000000000',
  `EquipmentHire` float(31,30) default '0.000000000000000000000000000000',
  `ExtendedServiceCost` float(31,30) default '0.000000000000000000000000000000',
  `CallCapping` float(31,30) default '0.000000000000000000000000000000',
  `FlagFall` float(31,30) default '0.000000000000000000000000000000',
  `OrderQuanity` int(11) default '1',
  `Stock` int(11) default '0',
  `WarehouseQuanityA` int(11) default '0',
  `PickingQuanityA` int(11) default '0',
  `FetchCode` varchar(42) default NULL,
  PRIMARY KEY  (`RecID`),
  KEY `ServiceID` (`ServiceID`,`VendorID`,`RadiusID`,`VendorPartID`,`SubPartID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `popsmtp` */

drop table if exists `popsmtp`;

CREATE TABLE `popsmtp` (
  `ID` bigint(20) NOT NULL default '0',
  `ipaddr` varchar(16) NOT NULL default '',
  `logtime` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`ID`)
) ENGINE=MEMORY DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

/*Table structure for table `purchaseorder` */

drop table if exists `purchaseorder`;

CREATE TABLE `purchaseorder` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `VendorID` bigint(20) default '0',
  `acci_RecID` bigint(20) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `DateSent` datetime default '1899-12-31 12:00:00',
  `POValue` float default '0',
  `POGST` float default '0',
  `ShippingID` bigint(20) default '0',
  `Finalised` tinyint(4) default '0',
  `Cancelled` tinyint(4) default '0',
  `Fradulant` tinyint(4) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `radiusaccounts` */

drop table if exists `radiusaccounts`;

CREATE TABLE `radiusaccounts` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Username` varchar(50) default NULL,
  `Password` varchar(50) default NULL,
  `SessionsAllowed` int(11) default '1',
  `AutoActivateFlag` tinyint(4) default '0',
  `Activate` datetime default NULL,
  `Deactivate` datetime default NULL,
  `SessionTimeout` int(11) default '10800',
  `IdleTimeout` int(11) default '600',
  `acci_RecID` bigint(20) default '0',
  `ServiceType` varchar(50) default 'Framed-User',
  `FramedProtocol` varchar(50) default 'PPP',
  `Checked` tinyint(4) default NULL,
  `sfStartTime` datetime default NULL,
  `sfStopTime` datetime default NULL,
  `sfAliveTime` datetime default NULL,
  `sfCycle_Upload` bigint(20) default '0',
  `sfCycle_Download` bigint(20) default '0',
  `sfCycle_Mins` bigint(20) default '0',
  `Acct_Session_ID` varchar(50) default NULL,
  `ptRecID` bigint(20) default '0',
  `VirtualID` bigint(20) default NULL,
  `PrimaryDNS` varchar(255) default NULL,
  `SecondaryDNS` varchar(255) default '203.17.15.163',
  `DateCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `radiuspools` */

drop table if exists `radiuspools`;

CREATE TABLE `radiuspools` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Description` varchar(50) default NULL,
  `ufFTPServer` varchar(255) default NULL,
  `ufTagetDir` varchar(255) default NULL,
  `ufPort` varchar(10) default NULL,
  `ufFilename` varchar(255) default NULL,
  `ufUsername` varchar(50) default NULL,
  `ufPassword` varchar(50) default NULL,
  `rlFTPServer` varchar(255) default NULL,
  `rlTagetDir` varchar(255) default NULL,
  `rlPort` varchar(10) default NULL,
  `rlFilename` varchar(255) default NULL,
  `rlUsername` varchar(50) default NULL,
  `rlPassword` varchar(50) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `receipts` */

drop table if exists `receipts`;

CREATE TABLE `receipts` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ReceiptNo` bigint(20) default '0',
  `acci_RecID` bigint(20) default '0',
  `Description` varchar(255) default NULL,
  `PaymentType` varchar(50) default 'Credit',
  `SerialNumber` varchar(255) default NULL,
  `RefundID` bigint(20) default '0',
  `acciServicesID` bigint(20) default '0',
  `TraxrID` bigint(20) default '0',
  `InvoiceOutID` bigint(20) default '0',
  `InvoiceInID` bigint(20) default '0',
  `Paid` float(13,5) default '0.00000',
  `Refunded` float(13,5) default '0.00000',
  `StatementID` double(13,5) default '0.00000',
  `GSTRefunded` float(13,5) default '0.00000',
  `GSTPaid` float(13,5) default '0.00000',
  `OnlineSessionSysopID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `recidallocations` */

drop table if exists `recidallocations`;

CREATE TABLE `recidallocations` (
  `RecID` bigint(20) default '0',
  `tblName` char(128) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `recidplacement` */

drop table if exists `recidplacement`;

CREATE TABLE `recidplacement` (
  `CountID` bigint(20) NOT NULL auto_increment,
  `TableName` varchar(100) default NULL,
  `RecID` bigint(20) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`CountID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `refundtraxr` */

drop table if exists `refundtraxr`;

CREATE TABLE `refundtraxr` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `RefundSerial` varchar(50) default '0',
  `acci_RecID` bigint(20) default NULL,
  `Total` float(42,2) default '0.00',
  `RefundDate` datetime default NULL,
  `Finalised` tinyint(4) default '0',
  `Comment` text,
  `CreatedWhen` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `SysopID` bigint(20) default NULL,
  `VirtualID` bigint(20) default NULL,
  `AmountRefunded` float(31,30) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `servicesupdatestatus` */

drop table if exists `servicesupdatestatus`;

CREATE TABLE `servicesupdatestatus` (
  `ID` int(11) NOT NULL auto_increment,
  `Apache` enum('Y','N') NOT NULL default 'Y',
  `Domain` enum('Y','N') NOT NULL default 'Y',
  `IMPOPALOID` enum('Y','N') NOT NULL default 'Y',
  `UIDAssign` enum('Y','N') NOT NULL default 'Y',
  UNIQUE KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Each time daemon info is updated set the field to ''Y'' so scr';

/*Table structure for table `servicetypes` */

drop table if exists `servicetypes`;

CREATE TABLE `servicetypes` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `SysopID` bigint(20) default '1',
  `VirtualID` bigint(20) default '0',
  `VendorID` bigint(20) default '0',
  `UpgradePathID` bigint(20) default '0',
  `nIcon` int(8) default '0',
  `ServiceKey` varchar(50) default NULL,
  `Description` varchar(50) default NULL,
  `SubofRecID` bigint(20) default NULL,
  `ListOnRadius` tinyint(4) default '0',
  `HasUID` tinyint(4) default '0',
  `HasSysUID` tinyint(4) default '0',
  `BillImmediately` tinyint(4) default '0',
  `SharingLevel` int(8) default '0',
  `MyResellerHierachalTree` tinyint(11) default '-1',
  `BreifDescription` varchar(255) default NULL,
  `LongDescription` text,
  `URL` text,
  `SecurityLevel` int(4) default '11',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `servicetypes_matrix` */

drop table if exists `servicetypes_matrix`;

CREATE TABLE `servicetypes_matrix` (
  `ServiceID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `SubServiceID` bigint(20) default '0'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `servicetypes_modekeys` */

drop table if exists `servicetypes_modekeys`;

CREATE TABLE `servicetypes_modekeys` (
  `RecID` int(11) NOT NULL auto_increment,
  `ServiceKey` varchar(50) default 'SALES',
  `GenericKey` varchar(50) default 'SALES',
  `showTempSQL` text,
  `showTempTitle` varchar(255) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `statementitems` */

drop table if exists `statementitems`;

CREATE TABLE `statementitems` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvRecID` bigint(20) default '0',
  `Items` tinyint(4) default '0',
  `Description` varchar(255) default NULL,
  `TotalDue` float(31,30) default '0.000000000000000000000000000000',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `stationary_codehtmlkeys` */

drop table if exists `stationary_codehtmlkeys`;

CREATE TABLE `stationary_codehtmlkeys` (
  `RecID` int(4) unsigned NOT NULL auto_increment,
  `StationaryCode` varchar(64) default NULL,
  `HTMLKey` text,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `superkeysindex` */

drop table if exists `superkeysindex`;

CREATE TABLE `superkeysindex` (
  `RecID` int(10) NOT NULL auto_increment,
  `sName` varchar(32) default NULL,
  `SuperKey` text,
  `RecUnlock` varchar(255) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `supplier` */

drop table if exists `supplier`;

CREATE TABLE `supplier` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `CompanyName` varchar(255) default NULL,
  `ACN` varchar(20) default NULL,
  `ABN` varchar(20) default NULL,
  `ContactName` varchar(50) default NULL,
  `Phone` varchar(30) default NULL,
  `Fax` varchar(30) default NULL,
  `ContactEmail` varchar(255) default NULL,
  `PurchaseOrderEmail` varchar(255) default NULL,
  `Comment` text,
  `NoItems` int(11) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `sysops` */

drop table if exists `sysops`;

CREATE TABLE `sysops` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `DateCreated` datetime default '2004-10-11 06:15:00',
  `Username` varchar(50) default NULL,
  `Password` varchar(50) default NULL,
  `Description` varchar(255) default NULL,
  `Checked` int(11) default '-1',
  `SecurityLevel` tinyint(2) default '0',
  `VirtualID` bigint(20) default '0',
  `AgencyID` bigint(20) default '0',
  `Master` int(11) default '0',
  `bMaintain` tinyint(4) default '0',
  `bVISP` tinyint(4) NOT NULL default '0',
  `LastNetworkAdd` text NOT NULL,
  `IP` varchar(255) default NULL,
  `IncomeTax` float(31,30) default '10.000000000000000000000000000000',
  `TFN` varchar(64) default NULL,
  `SuperRate` float(31,30) default '5.000000000000000000000000000000',
  `CommRate` float(31,30) default '0.000000000000000000000000000000',
  `PerVISP` float default '0',
  `TotalVisp` tinyint(2) default '0',
  `RateOnVISP` float(31,30) default NULL,
  `prjAlphaVersion` varchar(11) default '0.0.0',
  `bVISPFiscal` tinyint(4) default '0',
  `CommClass` int(11) default '0',
  `wrkPercent` tinyint(4) default NULL,
  `msg` text,
  `bAgency` tinyint(4) default '0',
  `bCreateSysop` tinyint(4) default '0',
  `bPrimary` tinyint(4) default '0',
  `bTemplates` tinyint(4) default '0',
  `Firstname` varchar(100) default NULL,
  `Surname` varchar(100) default NULL,
  `Email` varchar(255) default NULL,
  `Email1` varchar(255) default NULL,
  `Email2` varchar(255) default NULL,
  `Home` varchar(30) default NULL,
  `Work` varchar(30) default NULL,
  `Mobile` varchar(30) default NULL,
  `bPayMethod` int(1) default NULL,
  `AccountNo` varchar(100) default NULL,
  `BSB` varchar(20) default NULL,
  `bPayNo` varchar(30) default NULL,
  `Street1` varchar(100) default NULL,
  `Street2` varchar(100) default NULL,
  `Suburb` varchar(30) default NULL,
  `Postcode` varchar(20) default NULL,
  `State` varchar(50) default NULL,
  `Country` varchar(30) default NULL,
  `bDeleted` tinyint(4) default '0',
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `NextCycle` datetime default '0000-00-00 00:00:00',
  `PreviousCycle` datetime default '0000-00-00 00:00:00',
  `cFilledApplication` tinyint(4) default '0',
  `cFilledTaxation` tinyint(4) default '0',
  `cMailedForms` tinyint(4) default '0',
  `cBreif1` tinyint(4) default '0',
  `cBreif2` tinyint(4) default '0',
  `cCompleted` tinyint(4) default '0',
  `bRecievables` tinyint(4) default '0',
  `bInvoice` tinyint(4) default '0',
  `bExpenditure` tinyint(4) default '0',
  `bHoldings` tinyint(4) default '0',
  `bComm` tinyint(4) default '0',
  `bRefund` tinyint(4) default '0',
  `bAddCust` tinyint(4) default '0',
  `bOwnership` tinyint(4) default '0',
  `bAccSettings` tinyint(4) default '0',
  `bVendors` tinyint(4) default '0',
  `PublicKey` varchar(255) default 'IRCEncryptionisthebest',
  `bWEBAccount` tinyint(4) default '-1',
  `homepage` tinytext,
  `icq` varchar(64) default NULL,
  `aol` varchar(255) default NULL,
  `yahoo` varchar(255) default NULL,
  `msn` varchar(255) default NULL,
  `location` varchar(255) default NULL,
  `occupation` varchar(255) default NULL,
  `interests` varchar(255) default NULL,
  `biography` text,
  `INClause` text,
  `StartForcast` datetime default '2004-04-01 00:00:00',
  `EndForcast` datetime default '0000-00-00 00:00:00',
  `phpsessionid` varchar(255) default NULL,
  `bConfirmed` tinyint(4) default '0',
  `ConfirmByDate` datetime default NULL,
  `ConfirmationCode` varchar(48) default NULL,
  `xoops_userID` int(8) default '-1',
  `SysopNetworkHostname` tinytext,
  PRIMARY KEY  (`RecID`),
  KEY `VirtualID` (`Checked`,`VirtualID`,`AgencyID`,`Master`,`bVISP`,`bVISPFiscal`,`bAgency`,`bCreateSysop`,`bPrimary`,`bTemplates`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `tax` */

drop table if exists `tax`;

CREATE TABLE `tax` (
  `RecID` int(11) NOT NULL auto_increment,
  `Code` varchar(32) default NULL,
  `Percentage` float(13,5) default NULL,
  `Description` varchar(50) default NULL,
  `Created` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `Modified` datetime default NULL,
  `iFlag` int(11) default '0',
  `Country` varchar(64) default 'AUS0001',
  `RangeMin` float(12,7) default '0.0000000',
  `RangeMax` float(12,7) default '0.0000000',
  `FlatRate` float(32,30) default '0.000000000000000000000000000000',
  `lGroup` int(7) default '1',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `upgrade` */

drop table if exists `upgrade`;

CREATE TABLE `upgrade` (
  `RecID` int(11) NOT NULL auto_increment,
  `Version` varchar(11) default NULL,
  `Filename` varchar(50) default NULL,
  `MSI` tinyint(4) default NULL,
  `TimeCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `revision` int(11) NOT NULL default '0',
  `minor` int(11) NOT NULL default '0',
  `major` int(11) NOT NULL default '0',
  `Server` varchar(255) default 'ftp.comcen.com.au',
  `Port` varchar(255) default '21',
  `Username` varchar(255) default 'ant2003',
  `password` varchar(255) default 'z0κ',
  `remotedir` varchar(255) default '/home/syd/ant2003/downloads/bin',
  `searchpattern` varchar(25) default 'MOS??2005.msi',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `vendors` */

drop table if exists `vendors`;

CREATE TABLE `vendors` (
  `RecID` int(11) unsigned NOT NULL default '0',
  `vName` varchar(128) default NULL,
  `ABN` varchar(20) default NULL,
  `ACN` varchar(20) default NULL,
  `TFN` varchar(20) default NULL,
  `TaxFree` tinyint(4) default NULL,
  `Director1` varchar(128) default NULL,
  `Director2` varchar(128) default NULL,
  `Director3` varchar(128) default NULL,
  `BPay` varchar(128) default NULL,
  `BSB` varchar(20) default NULL,
  `Account` varchar(20) default NULL,
  `AccountName` varchar(128) default NULL,
  `Active` tinyint(4) default '-1',
  `SupportEmail` varchar(255) default NULL,
  `SupportPhone` varchar(50) default NULL,
  `poemailaddy` varchar(255) default NULL,
  `DSLDomain` varchar(255) default NULL,
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `SendPO` tinyint(4) default '-1',
  `ShareWithInternal` tinyint(4) default '0',
  `ShareWithExternal` tinyint(4) default '0',
  `TaxExpenditureOnly` tinyint(4) default '0',
  `SecuirtyLevel` int(8) default '10',
  `CodeMe` varchar(128) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `vendors_addresses` */

drop table if exists `vendors_addresses`;

CREATE TABLE `vendors_addresses` (
  `RecID` int(11) unsigned NOT NULL default '0',
  `VendorID` int(11) default '0',
  `ContactName` char(50) default NULL,
  `Street1` char(100) default NULL,
  `Street2` char(100) default NULL,
  `Suburb` char(50) default NULL,
  `State` char(20) default NULL,
  `Country` char(100) default NULL,
  `Postcode` char(100) default NULL,
  `Checked` tinyint(4) default '-1',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `vendors_email` */

drop table if exists `vendors_email`;

CREATE TABLE `vendors_email` (
  `RecID` int(11) unsigned NOT NULL default '0',
  `ContactName` char(50) default NULL,
  `EmailAddress` char(255) default NULL,
  `URL` char(255) default NULL,
  `Checked` tinyint(4) default '-1',
  `VendorID` int(11) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `vendors_phone` */

drop table if exists `vendors_phone`;

CREATE TABLE `vendors_phone` (
  `RecID` int(11) unsigned NOT NULL default '0',
  `VendorID` int(11) default NULL,
  `ContactName` char(50) default NULL,
  `PhoneNumber` char(50) default NULL,
  `Extension` char(20) default NULL,
  `ShortNote` char(255) default NULL,
  `Checked` tinyint(4) default '-1',
  `VendorIID` int(11) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `virtualisp` */

drop table if exists `virtualisp`;

CREATE TABLE `virtualisp` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `AgencyID` bigint(20) default '0',
  `VirtualID` bigint(20) default '0',
  `Description` varchar(100) default NULL,
  `BriefDesc` text,
  `Realm` varchar(255) default NULL,
  `CreationDate` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ABN` varchar(50) default '87096867775',
  `Subscribed` int(11) default '100',
  `SysopID` bigint(20) default '0',
  `ACN` varchar(50) default '096867775',
  `NoSub` int(11) default '0',
  `CreatedBy_SysopID` bigint(20) default NULL,
  `NextCycle` datetime default NULL,
  `PreviousCycle` datetime default NULL,
  `Cycle_IntervalType` varchar(4) default 'm',
  `Cycle_IntervalLength` int(11) default '1',
  `JoiningFee` float default '0',
  `LogoURL` varchar(255) default NULL,
  `Icon` int(11) default '1',
  `Manager` text,
  `Manager_SysopID` bigint(20) default '0',
  `AssistanceManager_SysopID` bigint(20) default '0',
  `Comment` text,
  `MISCFee` bigint(20) default '0',
  `bTaxMode` tinyint(4) default '0',
  `cTaxCode` varchar(255) default 'GST',
  `cTaxCountry` varchar(255) default 'AUS0001',
  `cTaxExemptCode` varchar(255) default NULL,
  `ftpFileDBMode` int(2) default '0',
  `ftpHostName` varchar(255) default '202.172.123.25',
  `ftpProxy` varchar(255) default NULL,
  `ftpUsername` varchar(32) default 'daemon',
  `ftpPassword` varchar(128) default '&4)υ>©&3ραWƒ¶`±',
  `ftpPort` int(4) default '21',
  `ftpBasePath` varchar(255) default '$FileDB$',
  `ftpGroupingFolder` varchar(128) default 'temp',
  `ftpNumberofFolders` int(11) default '0',
  `ftpNumberofFiles` int(11) default '0',
  `ftpTotalMB` int(11) default '0',
  `ftpCostPerMB` float(16,6) default '0.000000',
  `ftpURLPath` varchar(255) default 'http://202.172.123.25/$FileDB$/',
  `ftpPingAlive` int(2) default '1',
  `ftpIEProxy` int(2) default '1',
  `NumSales_AdminFeePerSale` float(16,6) default '0.000000',
  `NumSales_CappingAt` int(11) default '0',
  `NumSales_Minimum` int(11) default '0',
  `NumSales_Maximum` int(11) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `virtualisp_extended` */

drop table if exists `virtualisp_extended`;

CREATE TABLE `virtualisp_extended` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `RegStep` tinyint(4) default '2',
  `VirtualID` bigint(20) default '0',
  `TypeOfBusiness` text,
  `Finance_ContactName` varchar(128) default NULL,
  `Finance_PhoneNumber` varchar(32) default NULL,
  `Finance_FaxNumber` varchar(32) default NULL,
  `Finance_Email` varchar(255) default NULL,
  `Sales_ContactName` varchar(128) default NULL,
  `Sales_PhoneNumber` varchar(32) default NULL,
  `Sales_FaxNumber` varchar(32) default NULL,
  `Sales_Email` varchar(255) default NULL,
  `Admin_ContactName` varchar(128) default NULL,
  `Admin_PhoneNumber` varchar(32) default NULL,
  `Admin_FaxNumber` varchar(32) default NULL,
  `Admin_Email` varchar(255) default NULL,
  `Support_ContactName` varchar(128) default NULL,
  `Support_PhoneNumber` varchar(32) default NULL,
  `Support_FaxNumber` varchar(32) default NULL,
  `Support_Email` varchar(255) default NULL,
  `TradingType` varchar(150) default NULL,
  `History_EstablishedFor` varchar(20) default NULL,
  `History_DateCurOwnership` datetime default NULL,
  `Financial_PublicBank_Designation` varchar(255) default NULL,
  `Financial_PublicBank_Account_Number` varchar(64) default NULL,
  `Financial_PublicBank_Account_AccountBSB` varchar(6) default NULL,
  `Financial_PublicBank_Account_SwiftCode` varchar(32) default NULL,
  `Financial_PublicBank_Account_Name` varchar(255) default NULL,
  `Financial_PublicBank_Account_Style` varchar(128) default NULL,
  `Financial_Accountant_Name` varchar(128) default NULL,
  `Financial_Accountant_PhoneNumber` varchar(32) default NULL,
  `Financial_Accountant_FaxNumber` varchar(32) default NULL,
  `Financial_Accountant_Email` varchar(255) default NULL,
  `Financial_CreditLimit` float(11,4) default NULL,
  `Financial_CreditLimit_CRC` varchar(32) default NULL,
  `Financial_GlobalCreditLimit` float(11,4) default NULL,
  `Financial_GlobalCreditLimit_CRC` varchar(32) default NULL,
  `Financial_PaidUpCaptial` float(11,4) default NULL,
  `Financial_PaidUpCaptial_CRC` varchar(32) default NULL,
  `Financial_MonthlyAccessFee` float(11,4) default NULL,
  `Financial_MonthlyAccessFee_CRC` varchar(32) default NULL,
  `Financial_CustomerBlock_Step` int(8) default '150',
  `Financial_CustomerBlock_AdditionalBlock` float(11,4) default '4.3000',
  `DirectorA_Name` varchar(130) default NULL,
  `DirectorA_Address` varchar(255) default NULL,
  `DirectorA_IDType` varchar(100) default NULL,
  `DirectorA_IDNumber` varchar(100) default NULL,
  `DirectorA_DOB` date default NULL,
  `DirectorB_Name` varchar(130) default NULL,
  `DirectorB_Address` varchar(255) default NULL,
  `DirectorB_IDType` varchar(100) default NULL,
  `DirectorB_IDNumber` varchar(100) default NULL,
  `DirectorB_DOB` date default NULL,
  `DirectorC_Name` varchar(130) default NULL,
  `DirectorC_Address` varchar(255) default NULL,
  `DirectorC_IDType` varchar(100) default NULL,
  `DirectorC_IDNumber` varchar(100) default NULL,
  `DirectorC_DOB` date default NULL,
  `References_SupplierA_Name` varchar(255) default NULL,
  `References_SupplierA_ContactName` varchar(128) default NULL,
  `References_SupplierA_PhoneNumber` varchar(32) default NULL,
  `References_SupplierA_Email` varchar(255) default NULL,
  `References_SupplierA_Address` varchar(255) default NULL,
  `References_SupplierB_Name` varchar(255) default NULL,
  `References_SupplierB_ContactName` varchar(128) default NULL,
  `References_SupplierB_PhoneNumber` varchar(32) default NULL,
  `References_SupplierB_Email` varchar(255) default NULL,
  `References_SupplierB_Address` varchar(255) default NULL,
  `References_SupplierC_Name` varchar(255) default NULL,
  `References_SupplierC_ContactName` varchar(128) default NULL,
  `References_SupplierC_PhoneNumber` varchar(32) default NULL,
  `References_SupplierC_Email` varchar(255) default NULL,
  `References_SupplierC_Address` varchar(255) default NULL,
  `References_Client_Name` varchar(255) default NULL,
  `References_Client_ContactName` varchar(128) default NULL,
  `References_Client_PhoneNumber` varchar(32) default NULL,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `virtualisp_icons` */

drop table if exists `virtualisp_icons`;

CREATE TABLE `virtualisp_icons` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `IconNum` int(8) default '1',
  `IconMD5` varchar(32) default NULL,
  `VirtualID` bigint(20) default '0',
  `IconBlog` tinyblob,
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_accountheader` */

drop table if exists `visp_accountheader`;

CREATE TABLE `visp_accountheader` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `VirtualID` bigint(20) default NULL,
  `InvoiceID` bigint(20) default NULL,
  `StartDate` datetime default NULL,
  `EndDate` datetime default NULL,
  `TotalIncome` float(12,6) default '0.000000',
  `TotalCost` float(12,6) default '0.000000',
  `Paid` tinyint(4) default '0',
  `ChequeNumber` varchar(50) default NULL,
  `Margin` float(12,6) default '0.000000',
  `visp_PaymentID` bigint(20) default '0',
  `Income_GrossPayments` float(12,6) default '0.000000',
  `Income_AssessableGovernment` float(12,6) default '0.000000',
  `Income_IndustryPayments` float(12,6) default '0.000000',
  `Income_OtherBusiness` float(12,6) default '0.000000',
  `Income_TotalOverall` float(12,6) default '0.000000',
  `Expenses_Superannuation` float(12,6) default '0.000000',
  `Expenses_Commission` float(12,6) default '0.000000',
  `Expenses_CostofSale` float(12,6) default '0.000000',
  `Expenses_BadDebts` float(12,6) default '0.000000',
  `Expenses_Lease` float(12,6) default '0.000000',
  `Expenses_Rent` float(12,6) default '0.000000',
  `Expenses_TotalInterest` float(12,6) default '0.000000',
  `Expenses_TotalRoyalty` float(12,6) default '0.000000',
  `Expenses_Depreciation` float(12,6) default '0.000000',
  `Expenses_MotorVechicle` float(12,6) default '0.000000',
  `Expenses_RepairsMaintenance` float(12,6) default '0.000000',
  `Expenses_AllOther` float(12,6) default '0.000000',
  `Expenses_TotalOverall` float(12,6) default '0.000000',
  `Adjustments_ReconciliationItems` float(12,6) default '0.000000',
  `Adjustments_Income` float(12,6) default '0.000000',
  `Adjustments_Expenses` float(12,6) default '0.000000',
  `Adjustments_DroughtAllowance` float(12,6) default '0.000000',
  `sum_NetFromBusiness` float(12,6) default '0.000000',
  `FlagID` int(8) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_addresses` */

drop table if exists `visp_addresses`;

CREATE TABLE `visp_addresses` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `visp_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `ContactName` varchar(50) default NULL,
  `Street1` varchar(150) default NULL,
  `Street2` varchar(150) default NULL,
  `Country` varchar(50) default NULL,
  `State` varchar(50) default NULL,
  `Postcode` varchar(20) default NULL,
  `Suburb` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default '-1',
  `PhotoURL` tinytext,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_emailaddresses` */

drop table if exists `visp_emailaddresses`;

CREATE TABLE `visp_emailaddresses` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `visp_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `EmailAddress` varchar(255) default NULL,
  `ContactName` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default NULL,
  `PhotoURL` text,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_emailcatch` */

drop table if exists `visp_emailcatch`;

CREATE TABLE `visp_emailcatch` (
  `RecID` bigint(11) NOT NULL auto_increment,
  `VirtualID` bigint(20) default '0',
  `CatchType` varchar(48) default 'default',
  `CatchEmailAddress` varchar(255) default 'nexusarchives@ep.net.au',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_invoiceitems` */

drop table if exists `visp_invoiceitems`;

CREATE TABLE `visp_invoiceitems` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `InvoiceID` bigint(20) default NULL,
  `AmountDue` float(11,6) default '0.000000',
  `GSTPaid` float(11,6) default '0.000000',
  `CostPrice` float(11,6) default '0.000000',
  `GSTCharged` float(11,6) default '0.000000',
  `PaymentDue` datetime default NULL,
  `FinalClosureDate` datetime default NULL,
  `AmountPaid` float(11,6) default '0.000000',
  `PaidWhen` datetime default NULL,
  `TotalDue` float(11,6) default '0.000000',
  `TotalPaid` float(11,6) default NULL,
  `VirtualID` bigint(20) default '0',
  `Description` varchar(255) default NULL,
  `SysopID` bigint(20) default '0',
  PRIMARY KEY  (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_paymenthistory` */

drop table if exists `visp_paymenthistory`;

CREATE TABLE `visp_paymenthistory` (
  `VirtualID` bigint(20) default '0',
  `SysopID` bigint(20) default '0',
  `InvoiceID` bigint(20) default '0',
  `DateKeyed` datetime default NULL,
  `DateBanked` datetime default NULL,
  `RemittanceSent` int(2) default '0',
  `AmountPaid` float(12,6) default '0.000000',
  `GSTPaid` float(12,6) default '0.000000',
  `FlagID` int(8) default '0',
  `RecCreated` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Table structure for table `visp_phonenumbers` */

drop table if exists `visp_phonenumbers`;

CREATE TABLE `visp_phonenumbers` (
  `RecID` bigint(20) NOT NULL auto_increment,
  `visp_RecID` bigint(20) default '0',
  `FlagID` int(4) default '0',
  `DateAdded` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  `PhoneNumber` varchar(50) default NULL,
  `Extension` varchar(10) default NULL,
  `ContactName` varchar(50) default NULL,
  `Cancelled` int(11) default '0',
  `Checked` int(11) default NULL,
  `ShortNote` varchar(255) default NULL,
  `PhotoURL` text,
  PRIMARY KEY  (`RecID`),
  UNIQUE KEY `PrimaryKey` (`RecID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
