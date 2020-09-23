<?php

require("./DbSql.inc.php");

Class CustomSQL extends DBSQL
{
   // the constructor
   function CustomSQL($DBName = "")
   {
      $this->DBSQL($DBName);
   }
   
   function getalluser($page,$record)
   {
      $start = $page*$record;
      $sql = "select customerid,username from customer order by customerid DESC LIMIT $start,$record";
      $result = $this->select($sql);
      return $result;
   }
   
   function deluser($customerid)
   {            
      $sql = "DELETE FROM customer where customerid='$customerid'";
      $result = $this->delete($sql);      
      return $result;      
   }
   
   function adduser($username,$password,$email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography)
   {      
      $sql = "insert into customer (username,password,email,homepage,icq,aol,yahoo,msn,location,occupation,interests,biography) values ('$username','$password','$email','$homepage','$icq','$aol','$yahoo','$msn','$location','$occupation','$interests','$biography')";
      $result = $this->insert($sql);
      return $result;
   }
   
   function edituser($password,$email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$customerid)
   {
      $sql = "update customer set password='$password',email='$email',homepage='$homepage',icq='$icq',aol='$aol',yahoo='$yahoo',msn='$msn',location='$location',occupation='$occupation',interests='$interests',biography='$biography' where customerid='$customerid'";      
      $results = $this->update($sql);
      return $results;
   }
   
   function getuserinfobyid($customerid)
   {      
      $sql = "select * from customer where customerid='$customerid'";
      $result = $this->select($sql);
      return $result;
   }
   
   function getuserbykeyword($page,$record,$keyword)
   {
      $start = $page*$record;
      $sql = "select customerid,username from customer where username like '%$keyword%' order by customerid DESC LIMIT $start,$record";
      $result = $this->select($sql);
      return $result;
   }
      
}

?>