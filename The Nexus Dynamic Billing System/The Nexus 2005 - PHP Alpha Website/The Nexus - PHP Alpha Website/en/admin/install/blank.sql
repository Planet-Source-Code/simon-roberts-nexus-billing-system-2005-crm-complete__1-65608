CREATE TABLE customer (
   customerid bigint(32) NOT NULL auto_increment,
   username varchar(20),
   password varchar(20),
   email varchar(255),
   homepage varchar(255),
   icq varchar(50),
   aol varchar(50),
   yahoo varchar(50),
   msn varchar(100),
   location varchar(255),
   occupation varchar(255),
   interests varchar(255),
   biography varchar(255),
   PRIMARY KEY (customerid),
   UNIQUE username (username)
);

CREATE TABLE useradmin (
   adminid int(10) NOT NULL auto_increment,
   username varchar(50) NOT NULL,
   password varchar(50) NOT NULL,
   PRIMARY KEY (adminid)
);