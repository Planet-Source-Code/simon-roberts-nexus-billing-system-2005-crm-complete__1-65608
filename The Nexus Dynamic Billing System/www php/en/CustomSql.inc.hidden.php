
<?php

session_start();
require("DbSql.inc.hidden.php");

Class CustomSQL extends DBSQL
{
   // the constructor
   function CustomSQL($DBName = "")
   {
      $this->DBSQL($DBName);
   }

   function checkUsername($Username)
   {      
      $sql = "select RecID from sysops where Username='$Username'";
      $result = $this->select($sql);
      return $result;
   }  
   
   function adduser($Username,$Password,$Email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$Firstname,$Surname,$Description,$Checked,$VirtualID,$Home,$Work,$Mobile,$AccountNo,$BSB,$bPayNo,$Street1,$Street2,$Suburb,$Postcode,$State,$Country)
   {      
      $sql = "insert into sysops (INClause,bWEBAccount,Username,Password,Email,homepage,icq,aol,yahoo,msn,location,occupation,interests,biography,Firstname,Surname,Description,Checked,VirtualID,Home,Work,Mobile,AccountNo,BSB,bPayNo,Street1,Street2,Suburb,Postcode,State,Country) values ('(\'0\')','-1','$Username',encode('$Password','dr34mt1me'),'$Email','$homepage','$icq','$aol','$yahoo','$msn','$location','$occupation','$interests','$biography','$Firstname','$Surname','$Description','$Checked','$VirtualID','$Home','$Work','$Mobile','$AccountNo','$BSB','$bPayNo','$Street1','$Street2','$Suburb','$Postcode','$State','$Country')";
	  print $sql;
      $result = $this->insert($sql);
      return $result;
   }
   
   function logincheck($Username,$Password)
   {      
      $sql = "select * from sysops where Username='$Username' and decode(Password,'dr34mt1me')='$Password'";      
      $result = $this->select($sql);
      if (empty($result)) {
      return 0;
      }else{
      return $result;
      }
   }

	function mapvisps() {
	

	
	}
   
   function checkPassword($RecID,$Password)
   {      
      $sql = "select RecID from sysops where decode(Password,'dr34mt1me')='$Password' and RecID='$RecID'";
      $result = $this->select($sql);
      if (empty($result)) {
      return 0;
      }else{
      $CID = $result[0]["RecID"];
      return $CID;
      }
   }
   
   function Emailcheck($Email)
   {      
      $sql = "select decode(Password,'dr34mt1me') as Password from sysops where Email='$Email' or Email1='$Email' or Email2='$Email'";
      $result = $this->select($sql);
      if (empty($result)) {
      return 0;
      }else{
      $Password = $result[0]["Password"];
      return $Password;
      }
   }
   
   function getuserinfobyid($RecID)
   {      
      $sql = "select * from sysops where RecID='$RecID'";
      $result = $this->select($sql);
      return $result;
   }
   
   function edituser($Email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$RecID,$Firstname,$Surname,$Description,$Checked,$VirtualID,$Home,$Work,$Mobile,$AccountNo,$BSB,$bPayNo,$Street1,$Street2,$Suburb,$Postcode,$State,$Country)
   {
      $sql = "update sysops set Email='$Email',homepage='$homepage',icq='$icq',aol='$aol',yahoo='$yahoo',msn='$msn',location='$location',occupation='$occupation',interests='$interests',biography='$biography',Description='$Description',Checked='$Checked',VirtualID='$VirtualID',Home='$Home',Work='$Work',Mobile='$Mobile',AccountNo='$AccountNo',BSB='$BSB',bPayNo='$bPayNo',Street1='$Street1',Street2='$Street2',Suburb='$Suburb',Postcode='$Postcode',State='$State',Country='$Country',Firstname='$Firstname',Surname='$Surname' where RecID='$RecID'";      
      $results = $this->update($sql);
      return $results;
   }
   
   function modifypass($Password,$RecID)
   {
      $sql = "update sysops set Password=encode('$Password','dr34mt1me') where RecID='$RecID'";      
      $results = $this->update($sql);
      return $results;
   }
   
}

?>

