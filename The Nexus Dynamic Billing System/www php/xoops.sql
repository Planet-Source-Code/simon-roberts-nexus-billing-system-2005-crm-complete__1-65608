/* 
SQLyog v3.61
Host - demon.comcen.com.au : Database - epwebdev
**************************************************************
Server version 4.0.12-log
*/

create database if not exists `epwebdev`;

use `epwebdev`;

/*
Table struture for xoops_banner
*/

drop table if exists `xoops_banner`;
CREATE TABLE `xoops_banner` (
  `bid` int(4) NOT NULL auto_increment,
  `cid` int(2) NOT NULL default '0',
  `imptotal` int(8) NOT NULL default '0',
  `impmade` int(8) NOT NULL default '0',
  `clicks` int(8) NOT NULL default '0',
  `imageurl` varchar(255) NOT NULL default '',
  `clickurl` varchar(255) NOT NULL default '',
  `date` int(10) NOT NULL default '0',
  PRIMARY KEY  (`bid`),
  KEY `idxbannercid` (`cid`),
  KEY `idxbannerbidcid` (`bid`,`cid`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_banner
*/

INSERT INTO `xoops_banner` VALUES 
(2,2,20,1,0,'http://www.projectalpha.com.au/forum/images/banners/dc001.jpg','http://www.ep.net.au/',1089517536);

/*
Table struture for xoops_bannerclient
*/

drop table if exists `xoops_bannerclient`;
CREATE TABLE `xoops_bannerclient` (
  `cid` int(2) NOT NULL auto_increment,
  `name` varchar(60) NOT NULL default '',
  `contact` varchar(60) NOT NULL default '',
  `email` varchar(60) NOT NULL default '',
  `login` varchar(10) NOT NULL default '',
  `passwd` varchar(10) NOT NULL default '',
  `extrainfo` text NOT NULL,
  PRIMARY KEY  (`cid`),
  KEY `idxbannerclientlogin` (`login`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_bannerclient
*/

INSERT INTO `xoops_bannerclient` VALUES 
(2,'Dolphin Communication','Jarret Costi','jcosti@ep.net.au','jcosti','bringmetow','');

/*
Table struture for xoops_bannerfinish
*/

drop table if exists `xoops_bannerfinish`;
CREATE TABLE `xoops_bannerfinish` (
  `bid` int(4) NOT NULL auto_increment,
  `cid` int(2) NOT NULL default '0',
  `impressions` int(8) NOT NULL default '0',
  `clicks` int(8) NOT NULL default '0',
  `datestart` int(10) NOT NULL default '0',
  `dateend` int(10) NOT NULL default '0',
  PRIMARY KEY  (`bid`),
  KEY `idxbannerfinishcid` (`cid`)
) TYPE=MyISAM;

/*
Table struture for xoops_bb_categories
*/

drop table if exists `xoops_bb_categories`;
CREATE TABLE `xoops_bb_categories` (
  `cat_id` smallint(3) unsigned NOT NULL auto_increment,
  `cat_title` varchar(100) default NULL,
  `cat_order` varchar(10) default NULL,
  PRIMARY KEY  (`cat_id`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_bb_categories
*/

INSERT INTO `xoops_bb_categories` VALUES 
(1,'Project Alpha','1'),
(2,'Coding, Programming, Analyst and Design','2'),
(3,'ViSP Network Business Center','3');

/*
Table struture for xoops_bb_forum_access
*/

drop table if exists `xoops_bb_forum_access`;
CREATE TABLE `xoops_bb_forum_access` (
  `forum_id` int(4) unsigned NOT NULL default '0',
  `user_id` int(5) unsigned NOT NULL default '0',
  `can_post` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`forum_id`,`user_id`)
) TYPE=MyISAM;

/*
Table struture for xoops_bb_forum_mods
*/

drop table if exists `xoops_bb_forum_mods`;
CREATE TABLE `xoops_bb_forum_mods` (
  `forum_id` int(4) unsigned NOT NULL default '0',
  `user_id` int(5) unsigned NOT NULL default '0',
  KEY `forum_user_id` (`forum_id`,`user_id`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_bb_forum_mods
*/

INSERT INTO `xoops_bb_forum_mods` VALUES 
(1,1),
(2,1),
(3,1),
(4,1),
(5,1),
(6,1),
(7,1),
(8,1),
(9,1);

/*
Table struture for xoops_bb_forums
*/

drop table if exists `xoops_bb_forums`;
CREATE TABLE `xoops_bb_forums` (
  `forum_id` int(4) unsigned NOT NULL auto_increment,
  `forum_name` varchar(150) default NULL,
  `forum_desc` text,
  `forum_access` tinyint(2) NOT NULL default '1',
  `forum_moderator` int(2) default NULL,
  `forum_topics` int(8) NOT NULL default '0',
  `forum_posts` int(8) NOT NULL default '0',
  `forum_last_post_id` int(5) unsigned NOT NULL default '0',
  `cat_id` int(2) NOT NULL default '0',
  `forum_type` int(10) default '0',
  `allow_html` enum('0','1') NOT NULL default '0',
  `allow_sig` enum('0','1') NOT NULL default '0',
  `posts_per_page` tinyint(3) unsigned NOT NULL default '20',
  `hot_threshold` tinyint(3) unsigned NOT NULL default '10',
  `topics_per_page` tinyint(3) unsigned NOT NULL default '20',
  PRIMARY KEY  (`forum_id`),
  KEY `forum_last_post_id` (`forum_last_post_id`),
  KEY `cat_id` (`cat_id`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_bb_forums
*/

INSERT INTO `xoops_bb_forums` VALUES 
(1,'Requested Enchancements','This is where you can request a report or a feature to be added to project alpha, either on the website or the software.\r\n\r\nThis is of course so it can be discussed and properly analysied first time round. If you want another feature or you have a product that requires special handling in the routine please leave your request here so it can be actioned. \r\n\r\nSuch examples of feature are Trend Analyse Reports and Profit vs loss statements, BAS Statement',2,NULL,0,0,0,1,1,'1','1',10,10,20),
(2,'How do I?','This is for the question you have when wondering how to do a particular task or function.  This environment is dynamic so if you are unsure of a particular method of doing things. Please ask. This is so you can get support on a variety of topics and subject. \r\n\r\n',1,NULL,0,0,0,1,1,'1','1',10,10,20),
(3,'Bugs, Flaws and System Errors','This is where you report any errors or bugs you find in the software, this could be as simple as a hover selection that is causing problems to a minor hickup that may have happened allong the way. We rigiously test this software and use dynamic class based error checking so you will only see the error once. \r\n\r\nSo please we are not going to disbane your access for reporting a bug please feel free to do this at anytime.',1,NULL,0,0,0,1,0,'1','1',10,10,20),
(4,'General Enquires and Chat','If you have a question or would like to bring up a topic on anything from Mud Crabs to System Analyst and Design then you can do it here. \r\n\r\nPlease refrain from fowl language or abusive posts.',2,NULL,0,0,0,1,0,'1','1',10,10,20),
(5,'WAP, WML, XML','This is the point of chat in here is all on XML, WAP and WML. We have been experimenting with WML to bring some of project alphas features right to your very own Mobile Phone or PDA. This will mean there will be details available online via WAP.\r\n\r\nThese details include DSL Availability Checking, Change of User passwords, outstanding invoices. Answered Purchase orders.\r\n\r\nThis forum is also for code snippits so if you have any make sure you post them, we would love to help you debug.',2,NULL,0,0,0,2,0,'0','1',10,10,20),
(6,'Visual Basic 6.0 & .NET','This is the code snippit resource for visual basic. There are many sites on the web that will offer you source code for this language as VB has the most code of line ever written in any language on the planet today.\r\n\r\n',2,NULL,0,0,0,2,0,'0','1',10,10,20),
(7,'PHP, ASP & HTML','This is our humble resource on PHP and HTML. This programming language like ASP is the Unix variety of Dynamic Websites. These database driven sites allow the designer to have dynamic content and control. \r\n\r\nThis is still a new technology although similar to ASP, php can be driven from either a Windows Server or a Unix box. \r\n\r\nThis environment is what this website is based apon.',2,NULL,0,0,0,2,0,'0','1',10,10,20),
(8,'What is a ViSP? How do I get Started?','This is where you can all get introduced to the exciting word of ViSPing. A ViSP is our name for a Business that is using our network resources to conduct transactions and sales of products and services.\r\n\r\nA ViSP must have a ABN, ACN or RBN to conduct business on our network. This is within guideline set by the Australian Government and ASIC. \r\n\r\nSo you have all this, then this is a great place to start. Post a question and we will answer it to the best of our ability.  A business that uses project alpha must sign a Mutual Non Disclosure Agreement. This is to protect the privacy of both yourself, your staff and the clients you use to bill for subcribed services.\r\n\r\nWe offer you merchant and direct debit solutions as well. So going on what have you to loose, leave us a question and start your new enterprise on the digital infomation stream today.',2,NULL,0,0,0,3,0,'0','1',10,10,20),
(9,'New Product Announcements','So you have added a Template to the System, through one of your vendors and you want to tell other people on the network the great features of this product. \r\n\r\nPlease be as discriptive as possible and do not forget to put images and URL\'s to the products information pages. This is an ideal environment for you the Sysop to tell people of your new items details and to increase revenue coming back to you through other ViSP loading your template into their Sales Channel.',1,NULL,0,0,0,3,1,'0','1',10,10,20);

/*
Table struture for xoops_bb_posts
*/

drop table if exists `xoops_bb_posts`;
CREATE TABLE `xoops_bb_posts` (
  `post_id` int(8) unsigned NOT NULL auto_increment,
  `pid` int(8) NOT NULL default '0',
  `topic_id` int(8) NOT NULL default '0',
  `forum_id` int(4) NOT NULL default '0',
  `post_time` int(10) NOT NULL default '0',
  `uid` int(5) unsigned NOT NULL default '0',
  `poster_ip` varchar(15) default NULL,
  `subject` varchar(255) default NULL,
  `nohtml` tinyint(1) NOT NULL default '0',
  `nosmiley` tinyint(1) NOT NULL default '0',
  `icon` varchar(25) default NULL,
  `attachsig` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`post_id`),
  KEY `uid` (`uid`),
  KEY `pid` (`pid`),
  KEY `subject` (`subject`(40)),
  KEY `forumid_uid` (`forum_id`,`uid`),
  KEY `topicid_uid` (`topic_id`,`uid`),
  KEY `topicid_postid_pid` (`topic_id`,`post_id`,`pid`)
) TYPE=MyISAM;

/*
Table struture for xoops_bb_posts_text
*/

drop table if exists `xoops_bb_posts_text`;
CREATE TABLE `xoops_bb_posts_text` (
  `post_id` int(8) unsigned NOT NULL auto_increment,
  `post_text` text,
  PRIMARY KEY  (`post_id`)
) TYPE=MyISAM;

/*
Table struture for xoops_bb_topics
*/

drop table if exists `xoops_bb_topics`;
CREATE TABLE `xoops_bb_topics` (
  `topic_id` int(8) unsigned NOT NULL auto_increment,
  `topic_title` varchar(255) default NULL,
  `topic_poster` int(5) NOT NULL default '0',
  `topic_time` int(10) NOT NULL default '0',
  `topic_views` int(5) NOT NULL default '0',
  `topic_replies` int(4) NOT NULL default '0',
  `topic_last_post_id` int(8) unsigned NOT NULL default '0',
  `forum_id` int(4) NOT NULL default '0',
  `topic_status` tinyint(1) NOT NULL default '0',
  `topic_notify` tinyint(1) NOT NULL default '0',
  `topic_sticky` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`topic_id`),
  KEY `forum_id` (`forum_id`),
  KEY `topic_last_post_id` (`topic_last_post_id`),
  KEY `topic_poster` (`topic_poster`),
  KEY `topic_forum` (`topic_id`,`forum_id`),
  KEY `topic_sticky` (`topic_sticky`)
) TYPE=MyISAM;

/*
Table struture for xoops_comments
*/

drop table if exists `xoops_comments`;
CREATE TABLE `xoops_comments` (
  `comment_id` int(8) unsigned NOT NULL auto_increment,
  `pid` int(8) unsigned NOT NULL default '0',
  `item_id` int(8) unsigned NOT NULL default '0',
  `date` int(10) NOT NULL default '0',
  `user_id` int(5) NOT NULL default '0',
  `ip` varchar(15) NOT NULL default '',
  `subject` varchar(255) default NULL,
  `comment` text NOT NULL,
  `nohtml` tinyint(1) NOT NULL default '0',
  `nosmiley` tinyint(1) NOT NULL default '0',
  `noxcode` tinyint(1) NOT NULL default '0',
  `icon` varchar(25) NOT NULL default '',
  PRIMARY KEY  (`comment_id`),
  KEY `pid` (`pid`),
  KEY `item_id` (`item_id`),
  KEY `user_id` (`user_id`),
  KEY `subject` (`subject`(40))
) TYPE=MyISAM;

/*
Table struture for xoops_ephem
*/

drop table if exists `xoops_ephem`;
CREATE TABLE `xoops_ephem` (
  `eid` int(11) NOT NULL auto_increment,
  `did` int(2) NOT NULL default '0',
  `mid` int(2) NOT NULL default '0',
  `yid` int(4) NOT NULL default '0',
  `content` text NOT NULL,
  PRIMARY KEY  (`eid`),
  KEY `idxephemdidmid` (`did`,`mid`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_ephem
*/

INSERT INTO `xoops_ephem` VALUES 
(3,1,1,2000,'test');

/*
Table struture for xoops_groups
*/

drop table if exists `xoops_groups`;
CREATE TABLE `xoops_groups` (
  `groupid` int(5) NOT NULL auto_increment,
  `name` varchar(50) NOT NULL default '',
  `description` text,
  `type` varchar(10) NOT NULL default '',
  PRIMARY KEY  (`groupid`),
  KEY `type` (`type`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_groups
*/

INSERT INTO `xoops_groups` VALUES 
(1,'webmaster','webmasters of this site','Admin'),
(2,'General Chat','This is the general topics forum, from here you can raise issues on any subject. If you do not wish to register we have an unregsitered user section.','User'),
(3,'Anonymous Users','Anonymous Users Group','Anonymous'),
(4,'Announcements, News, Realease Notes','This is where we publish release information, technical notes to the running of project alpha and any relevant news. This is not a a public posting group. Please read the information within.','Admin'),
(5,'Bugs, Error Reports and Enchancement Requests','here is wher you can publish any bugs that you have come across in recent days of using any of our software. Remember if you don\'t report it, we can\'t fix it without your help so please. \r\n\r\nAlso if you have a enchancement you require such as a form for special handling of a product or service make sure you post a description of this requirement so we can commence analyst of the request. We will then comfirm with you the details and added to our software.','Custom'),
(6,'WAP, PHP, WML, HTML Coding Forum','Here is where we discuss issue raised with the Wireless Community regarding the implementation of WML protocol on our website. If you are a coder you are more than welcome to post in this forum. We welcome code snippits and helpful hints.','Admin'),
(7,'ViSP Network Business Center','This is the hub of the ViSP enterprise network. This scheme of software allows any registered business whether it is newly declared or an traditional family run business to contact business on the information superhighway. \r\n\r\nYour not a ViSP already! well that easy enquire with us or an existing ViSP to have your business sponsored into the system. This will provide you with powerful tools, the easy of credit and direct debit facilities and Online ordering and designated shipping of goods and services.\r\n\r\nCome in make an enquiry and start your Digitally  based system with our customisable software today. Simple registers on the site and download project alpha to get your mutual non disclosure agreement. \r\n\r\nThis Network is only open to registered businesses. That is a company or business that has a current and withstanding business number such as ABN, ACN or RBN.','Admin');

/*
Table struture for xoops_groups_blocks_link
*/

drop table if exists `xoops_groups_blocks_link`;
CREATE TABLE `xoops_groups_blocks_link` (
  `groupid` int(5) NOT NULL default '0',
  `block_id` int(5) NOT NULL default '0',
  `type` char(1) NOT NULL default ''
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_groups_blocks_link
*/

INSERT INTO `xoops_groups_blocks_link` VALUES 
(1,8,'R'),
(1,7,'R'),
(1,5,'R'),
(1,4,'R'),
(1,2,'R'),
(1,1,'R'),
(2,8,'R'),
(2,9,'R'),
(2,10,'R'),
(2,6,'R'),
(2,7,'R'),
(3,3,'R'),
(3,10,'R'),
(3,9,'R'),
(3,6,'R'),
(3,8,'R'),
(1,6,'R'),
(1,9,'R'),
(1,10,'R'),
(1,3,'R'),
(2,5,'R'),
(2,3,'R'),
(2,2,'R'),
(2,1,'R'),
(3,7,'R'),
(3,5,'R'),
(3,2,'R'),
(3,1,'R'),
(5,3,'R'),
(5,2,'R'),
(5,1,'R'),
(4,8,'R'),
(5,5,'R'),
(5,7,'R'),
(5,6,'R'),
(5,10,'R'),
(5,9,'R'),
(5,8,'R'),
(6,1,'R'),
(6,2,'R'),
(6,3,'R'),
(6,4,'R'),
(6,5,'R'),
(6,10,'R'),
(6,9,'R'),
(6,8,'R'),
(7,1,'R'),
(7,2,'R'),
(7,3,'R'),
(7,6,'R'),
(7,10,'R'),
(1,11,'R'),
(4,11,'R'),
(1,12,'R'),
(1,13,'R'),
(4,12,'R'),
(4,13,'R'),
(1,14,'R'),
(4,14,'R'),
(1,15,'R'),
(4,15,'R');

/*
Table struture for xoops_groups_modules_link
*/

drop table if exists `xoops_groups_modules_link`;
CREATE TABLE `xoops_groups_modules_link` (
  `groupid` int(5) NOT NULL default '0',
  `mid` int(5) NOT NULL default '0',
  `type` char(1) NOT NULL default '',
  KEY `groupid_type_mid` (`groupid`,`type`,`mid`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_groups_modules_link
*/

INSERT INTO `xoops_groups_modules_link` VALUES 
(1,1,'A'),
(1,2,'A'),
(1,3,'A'),
(1,4,'A'),
(1,5,'A'),
(1,6,'A'),
(1,1,'R'),
(1,2,'R'),
(1,3,'R'),
(1,4,'R'),
(1,5,'R'),
(1,6,'R'),
(2,1,'R'),
(2,2,'R'),
(3,1,'R'),
(3,2,'R'),
(4,1,'A'),
(4,2,'A'),
(4,3,'A'),
(4,4,'A'),
(4,5,'A'),
(4,6,'A'),
(4,1,'R'),
(4,2,'R'),
(4,3,'R'),
(4,4,'R'),
(4,5,'R'),
(4,6,'R'),
(5,1,'R'),
(5,2,'R'),
(6,2,'A'),
(6,1,'R'),
(6,2,'R'),
(7,2,'A'),
(7,1,'R'),
(7,2,'R');

/*
Table struture for xoops_groups_users_link
*/

drop table if exists `xoops_groups_users_link`;
CREATE TABLE `xoops_groups_users_link` (
  `groupid` int(5) NOT NULL default '0',
  `uid` int(5) NOT NULL default '0',
  KEY `groupid_uid` (`groupid`,`uid`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_groups_users_link
*/

INSERT INTO `xoops_groups_users_link` VALUES 
(1,1),
(2,1),
(2,3),
(4,1);

/*
Table struture for xoops_lastseen
*/

drop table if exists `xoops_lastseen`;
CREATE TABLE `xoops_lastseen` (
  `uid` int(5) unsigned NOT NULL default '0',
  `username` varchar(25) NOT NULL default '',
  `time` int(10) NOT NULL default '0',
  `ip` varchar(15) NOT NULL default '',
  `online` tinyint(1) unsigned NOT NULL default '0',
  KEY `username` (`username`),
  KEY `time` (`time`),
  KEY `uid` (`uid`),
  KEY `ip` (`ip`),
  KEY `online` (`online`)
) TYPE=MyISAM;

/*
Table struture for xoops_metafooter
*/

drop table if exists `xoops_metafooter`;
CREATE TABLE `xoops_metafooter` (
  `meta` text NOT NULL,
  `footer` text NOT NULL
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_metafooter
*/

INSERT INTO `xoops_metafooter` VALUES 
('XOOPS, Php, WAP, WML, Visual Basic, .NET, .OCX, Virtual Network, Virtual Online Game, project alpha, Exitstencil Press, Crass, Dead Kennedys, Black Flag, Dynamic, Reporting, Commission, Home Business, Coffee.','<a href=\"http://www.xoops.org/\"><img src=\"http://www.projectalpha.com.au/forum/images/s_poweredby.gif\" alt=\"XOOPS Official Website\" target=\"_blank\" /></a>\r\n\r\n');

/*
Table struture for xoops_modules
*/

drop table if exists `xoops_modules`;
CREATE TABLE `xoops_modules` (
  `mid` int(5) unsigned NOT NULL auto_increment,
  `name` varchar(150) NOT NULL default '',
  `version` int(5) NOT NULL default '100',
  `last_update` int(10) NOT NULL default '0',
  `weight` int(3) NOT NULL default '0',
  `isactive` tinyint(1) NOT NULL default '0',
  `dirname` varchar(25) NOT NULL default '',
  `hasmain` tinyint(1) NOT NULL default '0',
  `hasadmin` tinyint(1) NOT NULL default '0',
  `hassearch` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`mid`),
  KEY `hasmain` (`hasmain`),
  KEY `hasadmin` (`hasadmin`),
  KEY `hassearch` (`hassearch`),
  KEY `dirname` (`dirname`),
  KEY `name` (`name`(15))
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_modules
*/

INSERT INTO `xoops_modules` VALUES 
(1,'System Admin',100,1020877924,0,1,'system',0,1,0),
(2,'News',101,1020877924,1,1,'news',1,1,1),
(3,'Who\'s Online',100,1089505636,1,1,'whosonline',0,0,0),
(4,'Members',100,1089505636,1,1,'xoopsmembers',1,0,0),
(5,'Ephemerides',100,1089509971,1,1,'ephemerides',0,1,0),
(6,'Forum',100,1089509971,1,1,'newbb',1,1,1);

/*
Table struture for xoops_newblocks
*/

drop table if exists `xoops_newblocks`;
CREATE TABLE `xoops_newblocks` (
  `bid` int(5) unsigned NOT NULL auto_increment,
  `mid` int(5) unsigned NOT NULL default '0',
  `func_num` int(5) NOT NULL default '0',
  `options` varchar(255) NOT NULL default '',
  `name` varchar(150) NOT NULL default '',
  `position` tinyint(1) NOT NULL default '0',
  `title` varchar(150) NOT NULL default '',
  `content` text NOT NULL,
  `side` tinyint(1) NOT NULL default '0',
  `weight` int(5) NOT NULL default '0',
  `visible` tinyint(1) NOT NULL default '0',
  `type` char(1) NOT NULL default '',
  `c_type` char(1) NOT NULL default '',
  `isactive` tinyint(1) NOT NULL default '0',
  `dirname` varchar(50) NOT NULL default '',
  `func_file` varchar(50) NOT NULL default '',
  `show_func` varchar(50) NOT NULL default '',
  `edit_func` varchar(50) NOT NULL default '',
  PRIMARY KEY  (`bid`),
  KEY `mid` (`mid`),
  KEY `visible` (`visible`),
  KEY `side_visible_mid` (`side`,`visible`,`mid`),
  KEY `mid_funcnum` (`mid`,`func_num`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_newblocks
*/

INSERT INTO `xoops_newblocks` VALUES 
(1,1,1,'','User Block',1,'','',0,0,1,'S','',1,'system','system_blocks.php','b_system_user_show',''),
(2,1,2,'','Login Block',1,'','',0,0,1,'S','',1,'system','system_blocks.php','b_system_login_show',''),
(3,1,3,'','Search Block',1,'','',0,0,1,'S','',1,'system','system_blocks.php','b_system_search_show',''),
(4,1,4,'','Waiting Contents Block',1,'','',0,0,1,'S','',1,'system','system_blocks.php','b_system_waiting_show',''),
(5,1,5,'','Main Menu Block',0,'','',0,0,1,'S','',1,'system','system_blocks.php','b_system_main_show',''),
(6,1,6,'320|250|poweredby.gif','Site Info Block',0,'','',0,1,1,'S','H',1,'system','system_blocks.php','b_system_info_show','b_system_info_edit'),
(7,2,1,'','News Topics Block',0,'','',0,0,1,'M','',1,'news','news_topics.php','b_news_topics_show',''),
(8,2,2,'','Big Story Block',0,'','',1,0,1,'M','',1,'news','news_bigstory.php','b_news_bigstory_show',''),
(9,2,3,'counter|10','Top News Block',0,'','',4,0,1,'M','H',1,'news','news_top.php','b_news_top_show','b_news_top_edit'),
(10,2,4,'published|10','Recent News Block',0,'','',3,0,1,'M','H',1,'news','news_top.php','b_news_top_show','b_news_top_edit'),
(11,3,1,'1|10|20','Who\'s Online Block',0,'','',0,0,0,'M','',1,'whosonline','whosonline.php','b_whosonline_show','b_whosonline_edit'),
(12,4,1,'10|1','Top posters',0,'','',0,0,0,'M','',1,'xoopsmembers','members_posters.php','b_xoopsmembers_posters_show','b_xoopsmembers_posters_edit'),
(13,4,2,'10|1','New members',0,'','',0,0,0,'M','',1,'xoopsmembers','members_new.php','b_xoopsmembers_new_show','b_xoopsmembers_new_edit'),
(14,5,1,'','Ephemerides Block',0,'','',0,0,0,'M','',1,'ephemerides','ephemerides.php','b_ephemerides_show',''),
(15,6,1,'10|1','Recent Discussions in the Forums',0,'','',0,0,0,'M','',1,'newbb','newbb_new.php','b_newbb_new_show','b_newbb_new_edit');

/*
Table struture for xoops_priv_msgs
*/

drop table if exists `xoops_priv_msgs`;
CREATE TABLE `xoops_priv_msgs` (
  `msg_id` int(8) NOT NULL auto_increment,
  `msg_image` varchar(100) default NULL,
  `subject` varchar(100) default NULL,
  `from_userid` int(5) unsigned NOT NULL default '0',
  `to_userid` int(5) unsigned NOT NULL default '0',
  `msg_time` int(10) NOT NULL default '0',
  `msg_text` text,
  `read_msg` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`msg_id`),
  KEY `to_userid` (`to_userid`),
  KEY `idxprivmsgstouseridreadmsg` (`to_userid`,`read_msg`),
  KEY `idxprivmsgsmsgidfromuserid` (`msg_id`,`from_userid`)
) TYPE=MyISAM;

/*
Table struture for xoops_ranks
*/

drop table if exists `xoops_ranks`;
CREATE TABLE `xoops_ranks` (
  `rank_id` int(5) NOT NULL auto_increment,
  `rank_title` varchar(50) NOT NULL default '',
  `rank_min` int(10) NOT NULL default '0',
  `rank_max` int(10) NOT NULL default '0',
  `rank_special` int(2) NOT NULL default '0',
  `rank_image` varchar(255) default NULL,
  PRIMARY KEY  (`rank_id`),
  KEY `rank_min` (`rank_min`),
  KEY `rank_max` (`rank_max`),
  KEY `idxranksrankminrankmaxranspecial` (`rank_min`,`rank_max`,`rank_special`),
  KEY `idxranksrankspecial` (`rank_special`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_ranks
*/

INSERT INTO `xoops_ranks` VALUES 
(2,'Just popping in',0,20,0,''),
(3,'Not too shy to talk',21,40,0,'star.gif'),
(4,'Quite a regular',41,70,0,'stars2.gif'),
(5,'Just can\'t stay away',71,150,0,'stars3.gif'),
(6,'Home away from home',151,10000,0,'stars4.gif'),
(7,'Webmaster',-1,-1,1,'redstars5.gif'),
(8,'Moderator',-1,-1,1,'stars5.gif');

/*
Table struture for xoops_session
*/

drop table if exists `xoops_session`;
CREATE TABLE `xoops_session` (
  `uid` int(5) unsigned NOT NULL default '0',
  `time` int(10) NOT NULL default '0',
  `ip` varchar(15) NOT NULL default '',
  `hash` varchar(32) NOT NULL default '',
  KEY `uid` (`uid`),
  KEY `time` (`time`),
  KEY `hash_ip` (`hash`,`ip`),
  KEY `uid_ip` (`uid`,`ip`)
) TYPE=MyISAM;

/*
Table struture for xoops_smiles
*/

drop table if exists `xoops_smiles`;
CREATE TABLE `xoops_smiles` (
  `id` int(10) NOT NULL auto_increment,
  `code` varchar(50) default NULL,
  `smile_url` varchar(100) default NULL,
  `emotion` varchar(75) default NULL,
  PRIMARY KEY  (`id`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_smiles
*/

INSERT INTO `xoops_smiles` VALUES 
(1,':-D','icon_biggrin.gif','Very Happy'),
(2,':-)','icon_smile.gif','Smile'),
(3,':-(','icon_frown.gif','Sad'),
(4,':-o','icon_eek.gif','Surprised'),
(5,':-?','icon_confused.gif','Confused'),
(6,'8-)','icon_cool.gif','Cool'),
(7,':lol:','icon_lol.gif','Laughing'),
(8,':-x','icon_mad.gif','Mad'),
(9,':-P','icon_razz.gif','Razz'),
(10,':oops:','icon_redface.gif','Embaressed'),
(11,':cry:','icon_cry.gif','Crying (very sad)'),
(12,':evil:','icon_evil.gif','Evil or Very Mad'),
(13,':roll:','icon_rolleyes.gif','Rolling Eyes'),
(14,';-)','icon_wink.gif','Wink'),
(15,':pint:','icon_drink.gif','Another pint of beer'),
(16,':hammer:','icon_hammer.gif','ToolTimes at work'),
(17,':idea:','icon_idea.gif','I have an idea');

/*
Table struture for xoops_stories
*/

drop table if exists `xoops_stories`;
CREATE TABLE `xoops_stories` (
  `storyid` int(8) unsigned NOT NULL auto_increment,
  `uid` int(5) NOT NULL default '0',
  `title` varchar(255) default NULL,
  `created` int(10) NOT NULL default '0',
  `published` int(10) NOT NULL default '0',
  `hostname` varchar(20) NOT NULL default '',
  `nohtml` tinyint(1) NOT NULL default '0',
  `nosmiley` tinyint(1) NOT NULL default '0',
  `hometext` text NOT NULL,
  `bodytext` text NOT NULL,
  `counter` int(8) unsigned NOT NULL default '0',
  `topicid` int(8) unsigned NOT NULL default '1',
  `ihome` tinyint(1) NOT NULL default '0',
  `notifypub` tinyint(1) NOT NULL default '0',
  `type` varchar(10) NOT NULL default '0',
  `topicdisplay` tinyint(1) NOT NULL default '0',
  `topicalign` char(1) NOT NULL default 'R',
  PRIMARY KEY  (`storyid`),
  KEY `idxstoriestopic` (`topicid`),
  KEY `ihome` (`ihome`),
  KEY `uid` (`uid`),
  KEY `published_ihome` (`published`,`ihome`),
  KEY `title` (`title`(40)),
  KEY `created` (`created`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_stories
*/

INSERT INTO `xoops_stories` VALUES 
(2,1,'Welcome to the project alpha forum',1089509583,0,'202.172.123.25',1,0,'Welcome one and all to the new project alpha forum. Here we have a variety of topics you can choose from as well as some none related ones. I will be tieing the Sysop registration form into this one so if you register for the project alpha solution you will be able to access all the features of the forum as well.\r\n\r\nPlease feel free to post what you want when you want. We also have a section on WAP programming as we have been experimenting with WML and PHP in recent weeks. I hope if you have any questions, you ask them, as we will do our best to answer them.\r\n\r\n\r\nKind Regards\r\n\r\nSystem Adminsitration :-D ','',0,3,0,0,'user',1,'');

/*
Table struture for xoops_topics
*/

drop table if exists `xoops_topics`;
CREATE TABLE `xoops_topics` (
  `topic_id` int(4) unsigned NOT NULL auto_increment,
  `topic_pid` int(4) unsigned NOT NULL default '0',
  `topic_imgurl` varchar(20) NOT NULL default '',
  `topic_title` varchar(50) NOT NULL default '',
  PRIMARY KEY  (`topic_id`),
  KEY `pid` (`topic_pid`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_topics
*/

INSERT INTO `xoops_topics` VALUES 
(3,0,'W95MBX03.gif','Announcements And News'),
(4,0,'bug.gif','Bugs, Error Reports');

/*
Table struture for xoops_users
*/

drop table if exists `xoops_users`;
CREATE TABLE `xoops_users` (
  `uid` int(5) unsigned NOT NULL auto_increment,
  `name` varchar(60) NOT NULL default '',
  `uname` varchar(25) NOT NULL default '',
  `email` varchar(60) NOT NULL default '',
  `url` varchar(100) NOT NULL default '',
  `user_avatar` varchar(30) default NULL,
  `user_regdate` int(10) NOT NULL default '0',
  `user_icq` varchar(15) default NULL,
  `user_from` varchar(100) default NULL,
  `user_sig` text,
  `user_viewemail` tinyint(1) NOT NULL default '0',
  `actkey` varchar(8) default NULL,
  `user_aim` varchar(18) default NULL,
  `user_yim` varchar(25) default NULL,
  `user_msnm` varchar(25) default NULL,
  `pass` varchar(32) NOT NULL default '',
  `posts` int(8) default '0',
  `attachsig` tinyint(1) default '0',
  `rank` int(5) NOT NULL default '0',
  `level` int(5) NOT NULL default '1',
  `theme` varchar(100) NOT NULL default '',
  `timezone_offset` float(3,1) NOT NULL default '0.0',
  `last_login` int(10) NOT NULL default '0',
  `umode` varchar(10) NOT NULL default '',
  `uorder` tinyint(1) NOT NULL default '0',
  `user_occ` varchar(100) default NULL,
  `bio` tinytext NOT NULL,
  `user_intrest` varchar(150) default NULL,
  PRIMARY KEY  (`uid`),
  KEY `idxusersuname` (`uname`),
  KEY `idxusersemail` (`email`),
  KEY `idxusersuiduname` (`uid`,`uname`),
  KEY `idxusersunamepass` (`uname`,`pass`)
) TYPE=MyISAM;

/*
Table data for epwebdev.xoops_users
*/

INSERT INTO `xoops_users` VALUES 
(1,'','psyhkal','simon@projectalpaha.com.au','http://www.projectalpha.com.au/forum/','001.gif',1089500998,'','','',1,NULL,'','','','4cb1614a769c76a89529f277ff25428d',0,0,7,5,'phpkaox',0.0,1089509877,'flat',0,'','',''),
(2,'','sydney','lsd25@hotmail.com','http://www.ep.net.au/','228.gif',1089505748,'','','',0,'2d8ed7b3','','','','ce2752f9195cd2fd7f8e78737e8f968d',0,0,0,0,'',10.0,0,'thread',1,'','',''),
(3,'','simon','simon@projectalpha.com.au','http://www.ep.net.au/','228.gif',1089506125,'','','',0,'9855787f','','','','4cb1614a769c76a89529f277ff25428d',0,0,0,1,'',10.0,1089509633,'thread',1,'','','');

