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
$query_acciSQL = sprintf("SELECT accountinfo.RecID, accountinfo.AccountName, accountinfo.ActivationDate, accountinfo.ExpiryDate, accountinfo.DOB, accountinfo.sfCycle_Upload, accountinfo.sfCycle_Download, accountinfo.sfCycle_Mins, accountclass.Description as Class, virtualisp.Description as ViSP, virtualisp.LogoURL as Logo, sysops.Username FROM accountinfo, accountclass, virtualisp, sysops where accountinfo.Classification = accountclass.RecID and accountinfo.VirtualID = virtualisp.RecID and accountinfo.SysopID = sysops.RecID and accountinfo.RecID = %s", $accirecid);
$acciSQL = mysql_query($query_acciSQL, $projectalpha) or die(mysql_error());
$row_acciSQL = mysql_fetch_assoc($acciSQL);
$totalRows_acciSQL = mysql_num_rows($acciSQL);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset1 = sprintf("SELECT acci_phonenumbers.PhoneNumber, acci_phonenumbers.Extension, acci_phonenumbers.ContactName, acci_phonenumbers.ShortNote FROM acci_phonenumbers where acci_phonenumbers.AccI_RecID = %s ORDER BY acci_phonenumbers.DateAdded ", $accirecid);
$Recordset1 = mysql_query($query_Recordset1, $projectalpha) or die(mysql_error());

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset2 = sprintf("SELECT acci_addresses.ContactName, acci_addresses.Street1, acci_addresses.Street2, acci_addresses.Country, acci_addresses.`State`, acci_addresses.Postcode, acci_addresses.Suburb FROM acci_addresses where acci_addresses.AccI_RecID = %s  ORDER BY acci_addresses.DateCreated", $accirecid);
$Recordset2 = mysql_query($query_Recordset2, $projectalpha) or die(mysql_error());

mysql_select_db($database_projectalpha, $projectalpha);
$query_acciServ = sprintf("SELECT acci_services.RecID, acci_services.PeriodFee, acci_services.ContractExpiry, plantypes.Description, plantypes.CatNo FROM acci_services, plantypes where acci_services.ptRecID = plantypes.RecID and acci_services.AccI_RecID = %s", $accirecid);
$acciServ = mysql_query($query_acciServ, $projectalpha) or die(mysql_error());


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
<title>[ <?php echo $row_acciSQL['RecID']; ?> - <?php echo $row_acciSQL['AccountName']; ?> ] - [ <?php echo $row_acciSQL['Username']; ?> ]</title>
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
.style3 {font-size: 10px}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
.style10 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
}
.style11 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
}
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<table width="73%"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td colspan="5"><p>This report was generated for: <?php echo $row_acciSQL['Username']; ?></p>
      <p>&nbsp;</p></td>
  </tr>
  <tr>
    <td width="20%"><span class="style5"><?php echo $row_acciSQL['RecID']; ?></span></td>
    <td colspan="4"><span class="style5"><?php echo $row_acciSQL['AccountName']; ?></span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><span class="style8">Activation Date: </span></td>
    <td><span class="style8"><?php echo $row_acciSQL['ActivationDate']; ?></span></td>
    <td>&nbsp;</td>
    <td><span class="style8">Upstream Data:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['sfCycle_Upload']; ?></span></td>
  </tr>
  <tr>
    <td><span class="style8">Expiry Date:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['ExpiryDate']; ?></span></td>
    <td>&nbsp;</td>
    <td><span class="style8">Downstream Data:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['sfCycle_Download']; ?></span></td>
  </tr>
  <tr>
    <td><span class="style8">DOB or DOR: </span></td>
    <td><span class="style8"><?php echo $row_acciSQL['DOB']; ?></span></td>
    <td>&nbsp;</td>
    <td><span class="style8">Total Minutes On:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['sfCycle_Mins']; ?></span></td>
  </tr>
  <tr>
    <td><span class="style8">Classification:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['Class']; ?></span></td>
    <td>&nbsp;</td>
    <td><span class="style8">ViSP Network:</span></td>
    <td><span class="style8"><?php echo $row_acciSQL['ViSP']; ?></span></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5" class="style5"><div align="left">Phone Contacts for account</div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="29%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="23%">&nbsp;</td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
      <tr class="style8">
        <td><div align="center"><strong>Contact Name </strong></div></td>
        <td><div align="center"><strong>Phone Number </strong></div></td>
        <td><div align="center"><strong>Extension</strong></div></td>
        <td><div align="center"><strong>Short Note </strong></div></td>
      </tr>


       <?php 
      while ($row_Recordset1 = mysql_fetch_assoc($Recordset1)) {

  ?>
      <tr class="style8">
        <td><div align="center"><strong><?php echo $row_Recordset1['ContactName']; ?></strong></div></td>
        <td><div align="center"><strong><?php echo $row_Recordset1['PhoneNumber']; ?></strong></div></td>
        <td><div align="center"><strong><?php echo $row_Recordset1['Extension']; ?></strong></div></td>
        <td><div align="center"><strong><?php echo $row_Recordset1['ShortNote']; ?></strong></div></td>
      </tr>

	  <?php
	} ?>



      <tr class="style8">
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5" class="style5">Postal Contacts for account</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td width="9%">&nbsp;</td>
        <td width="82%"><div align="center" class="style8"><strong>Address In System </strong></div></td>
        <td width="9%">&nbsp;</td>
      </tr>

       <?php 
      while ($row_Recordset2 = mysql_fetch_assoc($Recordset2)) {

  ?>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center" class="style8">
          <p><strong><?php echo $row_Recordset2['ContactName']; ?></strong><br />
            <?php echo $row_Recordset2['Street1']; ?><br />
            <?php echo $row_Recordset2['Street2']; ?><br />
            <?php echo $row_Recordset2['State']; ?> <?php echo $row_Recordset2['Postcode']; ?>            <?php echo $row_Recordset2['Country']; ?><br />
          </p>
          </div></td>
        <td>&nbsp;</td>
      </tr>

	  <?php
	} ?>
	  


      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5" class="style5">Services for this account</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
      <tr class="style8">
        <td><div align="left"><strong>Product Code </strong></div></td>
        <td><div align="center"><strong>Description</strong></div></td>
        <td><div align="center"><strong>Cycle Period Fee <br />
          (Ex TAX)
        </strong></div></td>
        <td><div align="center"><strong>Contract Expires</strong></div></td>
      </tr>
       <?php 
      while ($row_acciServ = mysql_fetch_assoc($acciServ)) {

  ?>
	  <tr class="style8">
        <td><?php echo $row_acciServ['CatNo']; ?></td>
        <td><div align="center"><?php echo $row_acciServ['Description']; ?></div></td>
        <td><div align="center"><?php echo $row_acciServ['PeriodFee']; ?></div></td>
        <td><div align="center"><?php echo $row_acciServ['ContractExpiry']; ?></div></td>
      </tr>
	  <?php
	} ?>
	  
      <tr class="style8">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="center"><a href="login.php"><img src="../images/icons/cd-rom.jpg" width="32" height="32" /><br />
        <span class="style10">Back To Main</span> </a><br />
        <a href="clnt_explorer.php"><span class="style11">View your clients.</span></a> </div></td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>
<?php
mysql_free_result($acciSQL);

mysql_free_result($Recordset1);

mysql_free_result($Recordset2);

mysql_free_result($acciServ);



mysql_free_result($rsActive);
?>
