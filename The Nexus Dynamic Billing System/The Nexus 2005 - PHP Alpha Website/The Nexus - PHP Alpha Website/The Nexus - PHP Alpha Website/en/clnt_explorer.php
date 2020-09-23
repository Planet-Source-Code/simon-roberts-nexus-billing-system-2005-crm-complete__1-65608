<?php require_once('../connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}
if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view client base</a>";
	exit;
}

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsActive = sprintf("SELECT accountinfo.RecID, accountinfo.AccountName, accountinfo.BillingDate, accountclass.Description as Class, virtualisp.Description FROM accountinfo, accountclass, virtualisp where accountinfo.Classification = accountclass.RecID and accountinfo.VirtualID = virtualisp.RecID and accountinfo.Cancelled = 0 and accountinfo.SysopID = '%d'", $_SESSION['SysopID']);
$rsActive = mysql_query($query_rsActive, $projectalpha) or die(mysql_error());

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsCanned = sprintf("SELECT accountinfo.RecID, accountinfo.AccountName, accountinfo.BillingDate, accountclass.Description as Class, virtualisp.Description FROM accountinfo, accountclass, virtualisp where accountinfo.Classification = accountclass.RecID and accountinfo.VirtualID = virtualisp.RecID and accountinfo.Cancelled <> 0 and accountinfo.SysopID = '%d'", $_SESSION['SysopID']);
$rsCanned = mysql_query($query_rsCanned, $projectalpha) or die(mysql_error());

?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 5px;
	background-color: #DEDECA;
}
.style1 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #FF9933;
}
.style2 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style3 {font-size: 10px}
.style6 {
	font-size: 10px;
	font-weight: bold;
	color: #6666CC;
}
.style7 {color: #FF3333}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FF3333;
	font-weight: bold;
}
.style10 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
}
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<table width="73%"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td><h3 class="style1">Active Accounts On Your Client Base </h3></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><table width="99%"  border="0" cellspacing="1" cellpadding="0" align="center">
      <tr>
        <td width="30%"><div align="center" class="style3 style2 style7">
          <div align="left"><strong>Account Name </strong></div>
        </div></td>
        <td width="20%"><div align="center" class="style8">Billing Date </div></td>
        <td width="13%"><div align="center" class="style8">Num of Services </div></td>
        <td width="18%"><div align="center" class="style8">Class</div></td>
        <td width="19%"><div align="center" class="style8">ViSP</div></td>
      </tr>
	   <?php 
      while ($row_rsActive = mysql_fetch_assoc($rsActive)) {

		mysql_select_db($database_projectalpha, $projectalpha);
		$query_Recordset1 = sprintf("SELECT count(*) as NumServ FROM acci_services where AccI_RecID = %s", $row_rsActive['RecID']);
		$Recordset1 = mysql_query($query_Recordset1, $projectalpha) or die(mysql_error());
		$row_Recordset1 = mysql_fetch_assoc($Recordset1);
		
  ?>
      <tr>
        <td><div align="left"><span class="style2"><span class="style3"><a href="accidossier.php?accirecid=<?php echo $row_rsActive['RecID']; ?>"><?php echo $row_rsActive['AccountName']; ?></a></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsActive['BillingDate']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style6"><?php echo $row_Recordset1['NumServ']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsActive['Class']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsActive['Description']; ?></span></span></div></td>
      </tr>
	  <?php
	  mysql_free_result($Recordset1);
	} ?>
      <tr>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><span class="style1">Cancelled Accounts On Your Client Base </span></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><table width="99%"  border="0" cellspacing="1" cellpadding="0" align="center">
      <tr>
        <td width="30%"><div align="center" class="style3 style2 style7">
          <div align="left"><strong>Account Name </strong></div>
        </div></td>
        <td width="20%"><div align="center" class="style8">Billing Date </div></td>
        <td width="13%"><div align="center" class="style8">Num of Services </div></td>
        <td width="18%"><div align="center" class="style8">Class</div></td>
        <td width="19%"><div align="center" class="style8">ViSP</div></td>
      </tr>
	   <?php 
      while ($row_rsCanned = mysql_fetch_assoc($rsCanned)) {

		mysql_select_db($database_projectalpha, $projectalpha);
		$query_Recordset1 = sprintf("SELECT count(*) as NumServ FROM acci_services where AccI_RecID = %s", $row_rsCanned['RecID']);
		$Recordset1 = mysql_query($query_Recordset1, $projectalpha) or die(mysql_error());
		$row_Recordset1 = mysql_fetch_assoc($Recordset1);
		
  ?>
      <tr>
        <td><div align="left"><span class="style2"><span class="style3"><a href="accidossier.php?accirecid=<?php echo $row_rsCanned['RecID']; ?>"><?php echo $row_rsCanned['AccountName']; ?></a></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsCanned['BillingDate']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style6"><?php echo $row_Recordset1['NumServ']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsCanned['Class']; ?></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><?php echo $row_rsCanned['Description']; ?></span></span></div></td>
      </tr>
	  <?php
	  mysql_free_result($Recordset1);
	} ?>
      <tr>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"></span></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="center"><a href="login.php"><img src="../images/icons/cd-rom.jpg" width="32" height="32" /><br />
      <span class="style10">Back To Main</span> </a> </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>
<?php


mysql_free_result($rsActive);
?>
