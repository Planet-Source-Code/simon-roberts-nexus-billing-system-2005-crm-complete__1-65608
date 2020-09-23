<?php 
if (empty($nYSID) || empty($nSysopID) || empty($nMD5A) || empty($Email)) {

	echo "I am sorry you do not have access to this report at this time.";
	exit;
	
}
if(!session_id()){
  session_start();
}
 ?>
<?php require('../Connections/projectalpha.php'); ?>
<?php require('classes.php'); ?>
<?php
if (!(session_is_registered("step"))) {
	$_SESSION['step']=1;
} else {

	if ($_SESSION['step']==1) {
	
		if ($allowaccess==-1) {
			$_SESSION['step']=2;
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_prim = "update sysops set bWEBAccount='0'  where RecID = $nSysopID";
			$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
		} else {
			$_SESSION['step']=1;
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_prim = "update sysops set bWEBAccount='-1', bConfirmed='-1', VirtualID='0'  where RecID = $nSysopID";
			$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
		}
	} else {
			if ($_SESSION['step']==2) {
			
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_prim = "update sysops set bVISP = '$bVendors', bCreateSysop = '$bCreateSysop', bTemplates = '$bTemplates', bRecievables = '$bRecievables', bInvoice = '$bInvoice', bExpenditure = '$bExpenditure', bHoldings = '$bHoldings', bComm = '$bComm', bRefund = '$bRefund', bAddCust = '$bAddCust', bOwnership = '$bOwnership', bAccSettings = '$bAccSettings', bVendors ='$bVendors'  where RecID = $nSysopID";
				$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
				$_SESSION['step']=1;
			}
	}
}




?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Set Sysops Access to the Software</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	color: #66FFCC;
}
body {
	background-color: #9999CC;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 5px;
}
.style1 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	font-size: large;
	color: #666666;
}
.style2 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	font-size: small;
	color: #333333;
}
.style5 {font-family: "Trebuchet MS", Tahoma, Arial}
.style6 {color: #000000}
.style7 {font-family: "Trebuchet MS", Tahoma, Arial; color: #000000; }
.style11 {font-size: x-small}
.style12 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	font-size: x-small;
	font-weight: bold;
}
.style13 {font-size: medium}
.style14 {
	font-size: x-large;
	font-family: "Trebuchet MS", Tahoma, Arial;
}
-->
</style></head>

<body>
<?php
include("top.php3");

	mysql_select_db($database_projectalpha, $projectalpha);
	$query_prim = "select md5(decode(`Password`,'dr34mt1me')) as MD5A, bPrimary from sysops where RecID = $nYSID";
	$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
	$row_rsprim = mysql_fetch_assoc($rsprim);
	$totalRows_rsprim = mysql_num_rows($rsprim);
	
	if ($row_rsprim['bPrimary']==0 || $row_rsprim['MD5A']<>$nMD5A) {
	?>
		<div align="center" class="style14">The security on this link as changed it is no longer valid.
	      <?php
		exit;
	} else {
	
	
		
?>
        </div>
<table width="770"  border="0" align="center" cellpadding="0" cellspacing="0">

		
		<?php
			if ($_SESSION['step']<>2) {
				$_SESSION['step']=1;
				
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_sysop = "select sysops.Firstname, sysops.Username, sysops.RecID, sysops.VirtualID, virtualisp.Description as VispDesc from sysops inner join virtualisp on sysops.VirtualID = virtualisp.RecID and sysops.RecID = $nSysopID";
				$rssysop = mysql_query($query_sysop, $projectalpha) or die(mysql_error());
				$row_sysop = mysql_fetch_assoc($rssysop);
				$totalRows_sysop = mysql_num_rows($rssysop);
		
		?> <tr bgcolor="#CCCCCC">
    <td height="37" colspan="2"><strong><span class="style1">Authorise and Edit Sysop Permissions </span></strong></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
		<form name="form1" method="post" action="">
      <tr>
        <td height="100" colspan="2">
          <div align="left">
            <blockquote>
              <p class="style2"><span class="style1">Use these quick functions to turn on and off this sysops access to your tier in the software. Simply Click on the button to activate or deactivate the sysops access to the software available on this site.</span><br>
                <br>
                Sysop Username: <?php echo $row_sysop['Username'] ?><br>
                Sysop Firstname: <?php echo $row_sysop['Firstname'] ?><br>
                Sysop Identifier Number: <?php echo $row_sysop['RecID'] ?><br>
                <br>
                Current Nominated BNP: <?php echo $row_sysop['VispDesc'] ?><br> 
                <br>
              </p>
            </blockquote>
          </div></td>
        </tr>
      <tr align="center" valign="middle">
        <td width="4%"><div align="center" class="style5 style6">
          <p>&nbsp;      </p>
          <p>&nbsp;</p>
        </div></td>
        <td width="96%">
          <div align="left" class="style7">
            <p><strong>
<input name="allowaccess" type="radio" value="-1">              
Allow this sysop to access the nexus moment of sales terminal and the NSCE Admin.</strong></p>
            <p><strong>
              <input name="noaccess" type="radio" value="-1">
              Do not allow access to any software </strong></p>
          </div></td>
      </tr>
      <tr>
        <td height="75" colspan="2"><div align="center">
          <input name="Submit" type="submit" class="style1" value="Change Access to the Software.">
            </div>          
          <div align="center"></div></td>
        </tr>
	  </form>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td><?php 
  }
  if ($_SESSION['step']==2) {
  ?>
  </tr>
  <tr>
    <td height="38" colspan="2" bgcolor="#CCCCCC" class="style1"><strong>Edit Access to the Software and Levels of Software. </strong></td>
  </tr>
  <tr>
    <td height="486" colspan="2" bgcolor="#009999">
	<form action="" method="post" name="form2" class="style12">
	<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="124" colspan="3"><blockquote>
          <p><span class="style13">Change the settings below to control the level of access this sysop has to the software and your information. This is crutial that you select the correct options for this sysop.</span></p>
          </blockquote></td>
        </tr>
      <tr>
        <td width="34%" height="43">
              <div align="left">
    <input name="bVISP" type="checkbox" id="bVISP" value="-1">
  Can Create Resellers.     
              </div>
        </td>
        <td width="33%"><div align="left">
          <p><span class="style12">
                <input name="bVISPFiscal" type="checkbox" id="bVISPFiscal" value="-1">
                Can View Extended Reseller Commision Reports. </span></p>
        </div></td>
        <td width="33%"><div align="left">
          <p><span class="style12">
                <input name="bCreateSysop" type="checkbox" id="bCreateSysop" value="-1">
                Can Create Sysops. </span></p>
        </div></td>
      </tr>
      <tr>
        <td height="46"><div align="left">
          <p><span class="style12">
                <input name="bTemplates" type="checkbox" id="bTemplates" value="-1">
                Template Creation.</span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bRecievables" type="checkbox" id="bRecievables" value="-1">
                Access Recievables. </span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bInvoice" type="checkbox" id="bInvoice" value="-1">
                Access Invoicing and Invoicing Register. </span></p>
        </div></td>
      </tr>
      <tr>
        <td height="48"><div align="left">
          <p><span class="style12">
                <input name="bExpenditure" type="checkbox" id="bExpenditure" value="-1">
                Access To Expenditure Systems </span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bHoldings" type="checkbox" id="bHoldings" value="-1">
                Access to Client Base and Details.</span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bComm" type="checkbox" id="bComm" value="-1">
                Access to the Commission Calculator. </span></p>
        </div></td>
      </tr>
      <tr>
        <td height="50"><div align="left">
          <p><span class="style12">
                <input name="bRefund" type="checkbox" id="bRefund" value="-1">
                Access to the Refund Terminal. </span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bAddCust" type="checkbox" id="bAddCust" value="-1">
                Able to add a new client.</span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bOwnership" type="checkbox" id="bOwnership" value="-1">
                Can change Ownership. </span></p>
        </div></td>
      </tr>
      <tr>
        <td height="44"><div align="left">
          <p><span class="style12">
                <input name="bAccSettings" type="checkbox" id="bAccSettings" value="-1">
                Able to import your templates as sales items. </span></p>
        </div></td>
        <td><div align="left">
          <p><span class="style12">
                <input name="bVendors" type="checkbox" id="bVendors" value="-1">
                Able to Add and view vendors. </span></p>
        </div></td>
        <td><p>&nbsp;</p>          <div align="left"><span class="style11"></span></div></td>
      </tr>
      <tr>
        <td height="55">&nbsp;</td>
        <td><div align="center">
          <input name="Submit2" type="submit" class="style1" value="Change Permissions And Go Back">
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
	</form></td>
  </tr>
  <tr bgcolor="#333333">
    <td width="387" height="32">&nbsp;</td>
    <td width="383">&nbsp;</td>
  </tr>
  <?php
  }
 ?>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
 <?php
  }
 ?>
</body>
</html>
