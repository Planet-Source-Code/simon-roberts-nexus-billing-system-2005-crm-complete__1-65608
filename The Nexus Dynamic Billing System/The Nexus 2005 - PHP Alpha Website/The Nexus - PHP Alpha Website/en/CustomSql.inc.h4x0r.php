
<?php

require("DbSql.inc.h4x0r.php");

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
   
   function adduser($sessionid, $Username,$Password,$Email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$Firstname,$Surname,$Description,$Checked,$VirtualID,$Home,$Work,$Mobile,$AccountNo,$BSB,$bPayNo,$Street1,$Street2,$Suburb,$Postcode,$State,$Country)
   {    
  
   	  $level = 1; 
	  $rank = 0;
	  $fullname = sprintf('%s %s', $Firstname, $Surname);
	  
      $sql = "insert into epwebdev.xoops_2004_users (name, uname, email, url, user_icq, user_aim, user_yim, user_msnm, pass, level, rank, user_from, user_occ, bio, user_intrest) Values ('$fullname','$Username','$Email','$homepage','$icq','$aol','$yahoo','$msn',md5('$Password'),'$level','$rank','$location','$occupation','$biography','$interests')";
	  $resultxoops = $this->insertprimary($sql); 
	  $sql = "insert into sysops (phpsessionid, ConfirmationCode, ConfirmByDate, xoops_userID, DateCreated,INClause,bWEBAccount,Username,Password,Email,homepage,icq,aol,yahoo,msn,location,occupation,interests,biography,Firstname,Surname,Description,Checked,VirtualID,Home,Work,Mobile,AccountNo,BSB,bPayNo,Street1,Street2,Suburb,Postcode,State,Country) values ('$sessionid',md5(DATE_ADD(NOW(), INTERVAL 11 DAY)),DATE_ADD(NOW(), INTERVAL 12 DAY), '$resultxoops',NOW(),'(\'0\')','-1','$Username',encode('$Password','dr34mt1me'),'$Email','$homepage','$icq','$aol','$yahoo','$msn','$location','$occupation','$interests','$biography','$Firstname','$Surname','$Description','$Checked','$VirtualID','$Home','$Work','$Mobile','$AccountNo','$BSB','$bPayNo','$Street1','$Street2','$Suburb','$Postcode','$State','$Country')";
	  $result = $this->insert($sql);
      return $result;
   }
   
   function logincheck($Username,$Passwordmd5,$sessionid)
   {      
	
      $sql = "select * from sysops where phpsessionid='$sessionid'";      
	  $result = $this->select($sql);
      if (empty($result)) {
		  $sql = "select * from sysops where Username='$Username' and md5(decode(`Password`,'dr34mt1me'))='$Passwordmd5'";      
		  $result = $this->select($sql);
		  if (empty($result)) {
			  return 0;
		  }else{
		  	  $sql = sprintf("update sysops set phpsessionid='%s', SysopNetworkHostname='%s' where RecID = %d",$sessionid,gethostbyaddr($REMOTE_ADDR),$result[0]['RecID']);      
			   $resultb = $this->update($sql);
			  return $result;
		  }
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
   
   function emailcheck($Email)
   {      
	$hostname_projectalpha = "demon.comcen.com.au";
	$database_projectalpha = "projectalpha";
	$username_projectalpha = "sroberts";
	$password_projectalpha = "kl7jb0lsf";
	$projectalpha = mysql_pconnect($hostname_projectalpha, $username_projectalpha, $password_projectalpha) or trigger_error(mysql_error(),E_USER_ERROR); 
	$sql = sprintf("select `Username`, decode(`Password`,'dr34mt1me') as decPassword from sysops where Email='%s' or Email1='%s' or Email2='%s'",$Email,$Email,$Email);
	mysql_select_db($database_projectalpha, $projectalpha);
	$ProductCount = mysql_query($sql, $projectalpha) or die(mysql_error());
	$result = mysql_fetch_assoc($ProductCount);
	$totalRows_ProductCount = mysql_num_rows($ProductCount);
      if ($totalRows_ProductCount==0) {
	      return 0;
	  } else {
	   $Password = "
";
       $Password .= "Password: " ;
	   $Password .= $result["decPassword"];
	   $Password .= "
";
   	   $Password .= "Username: ";
   	   $Password .= $result["Username"];
	   $Password .= "
";
	   $Password .= "
";
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

