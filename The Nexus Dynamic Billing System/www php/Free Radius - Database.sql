/*
SQLyog Enterprise v4.06 RC1
Host - 5.0.21-community-nt : Database - radius
*********************************************************************
Server version : 5.0.21-community-nt
*/


create database if not exists `radius`;

USE `radius`;

/*Table structure for table `radiusalive` */

drop table if exists `radiusalive`;

CREATE TABLE `radiusalive` (
  `RadAcctId` varchar(19) default NULL,
  `AcctSessionId` varchar(32) default NULL,
  `AcctUniqueId` varchar(32) default NULL,
  `UserName` varchar(64) default NULL,
  `Realm` varchar(64) default NULL,
  `NASIPAddress` varchar(15) default NULL,
  `NASPortId` bigint(20) default NULL,
  `NASPortType` varchar(32) default NULL,
  `AcctStartTime` datetime default NULL,
  `AcctStopTime` datetime default NULL,
  `AcctSessionTime` bigint(20) default NULL,
  `AcctAuthentic` varchar(32) default NULL,
  `AcctStatusType` varchar(32) default NULL,
  `ConnectInfo_start` varchar(32) default NULL,
  `ConnectInfo_stop` varchar(32) default NULL,
  `AcctInputOctets` bigint(20) default NULL,
  `AcctOutputOctets` bigint(20) default NULL,
  `CalledStationId` varchar(10) default NULL,
  `CallingStationId` varchar(10) default NULL,
  `AcctTerminateCause` varchar(32) default NULL,
  `ServiceType` varchar(32) default NULL,
  `FramedProtocol` varchar(32) default NULL,
  `FramedIPAddress` varchar(15) default NULL,
  `LOGTime` datetime default NULL,
  `AcctStartDelay` bigint(20) default NULL,
  `AcctStopDelay` bigint(20) default NULL,
  `chkAcctInputOctets` bigint(20) default NULL,
  `chkAcctOutputOctets` bigint(20) default NULL,
  `chkAcctStartTime` datetime default NULL
);

/*Table structure for table `radiusgroupreply` */

drop table if exists `radiusgroupreply`;

CREATE TABLE `radiusgroupreply` (
  `id` varchar(19) default NULL,
  `groupname` varchar(128) default NULL,
  `attribute` varchar(255) default NULL,
  `value` varchar(255) default NULL,
  `op` varchar(50) default NULL
);

/*Table structure for table `radiuslog` */

drop table if exists `radiuslog`;

CREATE TABLE `radiuslog` (
  `RadAcctId` varchar(19) default NULL,
  `AcctSessionId` varchar(32) default NULL,
  `AcctUniqueId` varchar(32) default NULL,
  `UserName` varchar(64) default NULL,
  `Realm` varchar(64) default NULL,
  `NASIPAddress` varchar(15) default NULL,
  `NASPortId` bigint(20) default NULL,
  `NASPortType` varchar(32) default NULL,
  `AcctStartTime` datetime default NULL,
  `AcctStopTime` datetime default NULL,
  `AcctSessionTime` bigint(20) default NULL,
  `AcctAuthentic` varchar(32) default NULL,
  `AcctStatusType` varchar(32) default NULL,
  `ConnectInfo_start` varchar(32) default NULL,
  `ConnectInfo_stop` varchar(32) default NULL,
  `DeleteAlive` char(1) default NULL,
  `AcctInputOctets` bigint(20) default NULL,
  `AcctOutputOctets` bigint(20) default NULL,
  `CalledStationId` varchar(10) default NULL,
  `CallingStationId` varchar(10) default NULL,
  `AcctTerminateCause` varchar(32) default NULL,
  `ServiceType` varchar(32) default NULL,
  `FramedProtocol` varchar(32) default NULL,
  `FramedIPAddress` varchar(15) default NULL,
  `AcctStartDelay` bigint(20) default NULL,
  `AcctStopDelay` bigint(20) default NULL,
  `FlagID` bigint(20) default NULL,
  `LOGTime` datetime default NULL
);

/*Table structure for table `radiusradcheck` */

drop table if exists `radiusradcheck`;

CREATE TABLE `radiusradcheck` (
  `id` varchar(19) default NULL,
  `username` varchar(16) default NULL,
  `attribute` varchar(255) default NULL,
  `value` varchar(255) default NULL,
  `op` char(2) default NULL,
  `pool` varchar(50) default NULL,
  `RadiusID` varchar(19) default NULL
);

/*Table structure for table `radiusradgroupcheck` */

drop table if exists `radiusradgroupcheck`;

CREATE TABLE `radiusradgroupcheck` (
  `id` bigint(20) default NULL,
  `GroupName` varchar(128) default NULL,
  `Attribute` varchar(255) default NULL,
  `Value` varchar(255) default NULL,
  `Op` char(2) default NULL
);

/*Table structure for table `radiusradreply` */

drop table if exists `radiusradreply`;

CREATE TABLE `radiusradreply` (
  `id` varchar(19) default NULL,
  `username` varchar(16) default NULL,
  `attribute` varchar(255) default NULL,
  `value` varchar(255) default NULL,
  `op` varchar(50) default NULL
);

/*Table structure for table `radiususergroup` */

drop table if exists `radiususergroup`;

CREATE TABLE `radiususergroup` (
  `id` varchar(19) default NULL,
  `username` varchar(16) default NULL,
  `groupname` varchar(128) default NULL
);
