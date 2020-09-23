<?php require_once('Connections/Epwebdev.php'); ?>
<?php
mysql_select_db($database_Epwebdev, $Epwebdev);
$query_redirect = "SELECT rdFiles.Filename, rdFiles.URLPath FROM rdFiles where ID = $nFileID";
$redirect = mysql_query($query_redirect, $Epwebdev) or die(mysql_error());
$row_redirect = mysql_fetch_assoc($redirect);
$totalRows_redirect = mysql_num_rows($redirect);
?>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">


<title>Project Alpha Redirection</title>
<HEAD>
<META HTTP-EQUIV="refresh" CONTENT="9;URL=<?php echo sprintf("%s%s",$row_redirect['URLPath'],$row_redirect['Filename']); ?>">
</HEAD>
<body background="mages/TimeDepentendLinearSystem.jpg" tracingsrc="i" tracingopacity="24">

<div align="center">
  <p>
  
    
<center>
    <br>
    <font face="Trebuchet MS"><b>Thankyou for choosing this Fantasic<br />
    Ground Breaking Digital Technology to look at.<br />
    <br />
    </b></font><font face="Trebuchet MS"><b>
	            <A HREF='<?php echo $row_redirect['URLPath']; ?><?php echo $row_redirect['Filename']; ?>' TARGET='_blank'>Click here if your download does not start within 10 secs</A></b></font>
	            <BR>
	            <BR>
	            <SCRIPT LANGUAGE=javascript>
			setTimeout('document.all.cdown.innerHTML = 28;',2000);
			setTimeout('document.all.cdown.innerHTML = 27;',3000);
			setTimeout('document.all.cdown.innerHTML = 26;',4000);
			setTimeout('document.all.cdown.innerHTML = 25;',5000);
			setTimeout('document.all.cdown.innerHTML = 24;',6000);
			setTimeout('document.all.cdown.innerHTML = 23;',7000);
			setTimeout('document.all.cdown.innerHTML = 22;',8000);
			setTimeout('document.all.cdown.innerHTML = 21;',9000);
			setTimeout('document.all.cdown.innerHTML = 20;',10000);
			setTimeout('document.all.cdown.innerHTML = 19;',12000);
			setTimeout('document.all.cdown.innerHTML = 18;',13000);
			setTimeout('document.all.cdown.innerHTML = 17;',14000);
			setTimeout('document.all.cdown.innerHTML = 16;',15000);
			setTimeout('document.all.cdown.innerHTML = 15;',16000);
			setTimeout('document.all.cdown.innerHTML = 14;',17000);
			setTimeout('document.all.cdown.innerHTML = 13;',18000);
			setTimeout('document.all.cdown.innerHTML = 12;',19000);
			setTimeout('document.all.cdown.innerHTML = 11;',20000);
			setTimeout('document.all.cdown.innerHTML = 10;',21000);
			setTimeout('document.all.cdown.innerHTML = 9;',22000);
			setTimeout('document.all.cdown.innerHTML = 8;',23000);
			setTimeout('document.all.cdown.innerHTML = 7;',24000);
			setTimeout('document.all.cdown.innerHTML = 6;',25000);
			setTimeout('document.all.cdown.innerHTML = 5;',26000);
			setTimeout('document.all.cdown.innerHTML = 4;',27000);
			setTimeout('document.all.cdown.innerHTML = 3;',28000);
			setTimeout('document.all.cdown.innerHTML = 2;',29000);
			setTimeout('document.all.cdown.innerHTML = 1;',30000);
			setTimeout('opener.location.reload();this.close();',31000);
			    </SCRIPT>
  <font face=arial size=2 color=black>This window will close in <font color=red face=arial size=3><b><span id=cdown>29</span></b></font> seconds!</font></p>
</div>
</body>
</html>
<?php
mysql_free_result($redirect);
?>
