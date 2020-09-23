<?php 
if(!session_id()){
  session_start();
} ?>
<HTML>
<HEAD>
<title>The Nexus Registration for Sysop and Bonus Rankings in Forum</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="./style/style.css" type="text/css">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
<link href="../css/txtbox3.css" rel="stylesheet" type="text/css">
<link href="txtbox2.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style3 {font-weight: bold}
.style6 {
	font-size: large;
	color: #FF9933;
}
.style13 {
	font-size: small;
	font-style: italic;
	font-weight: bold;
	color: #FF9900;
}
.style19 {font-size: small; font-weight: bold; }
.style22 {
	font-size: 12px;
	font-weight: bold;
}
.style24 {
	font-weight: bold;
	font-size: medium;
	color: #FFFFFF;
}
.style25 {font-size: 9px}
.style26 {
	color: #CCCCFF;
	font-size: medium;
}
.style30 {color: #000099}
.style31 {font-family: "Trebuchet MS", Tahoma, Arial}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style33 {font-size: small; color: #FF9900; font-family: "Trebuchet MS", Tahoma, Arial; font-weight: bold; }
.style34 {font-style: italic; font-weight: bold; font-size: 24px; font-family: "Trebuchet MS", Tahoma, Arial;}
.style35 {color: #FF9900}
.style36 {font-size: small; font-weight: bold; color: #FF9900; }
.style38 {font-size: small; font-weight: bold; color: #000099; }
.style41 {color: #FF6600}
.style43 {
	color: #FFFFFF;
	font-weight: bold;
}
.style44 {font-size: medium}
.style45 {color: #FFCCCC}
.style47 {
	color: #FF0000;
	font-weight: bold;
}
.style48 {color: #FFFFFF}
-->
</style>
</head>

<?php require('../Connections/projectalpha.php'); ?>
<?php require('../Connections/Epwebdev.php'); ?>
<?php require('classes.php'); ?>

<?php

if (empty($ConfirmationCode) or empty($Email)) {

	mysql_select_db($database_projectalpha, $projectalpha);
	$query_rsFetchNewbie = "select phpsessionid, ConfirmByDate, bConfirmed, ConfirmationCode, xoops_userID, Email, Username, decode(`Password`,'dr34mt1me') as md5Pass, Description, Firstname, Surname from sysops where ConfirmByDate > NOW() and RecID = '$nSysopID'";
	$rsFetchNewbie = mysql_query($query_rsFetchNewbie, $projectalpha) or die(mysql_error());
	$row_rsFetchNewbie = mysql_fetch_assoc($rsFetchNewbie);
	$totalRows_rsFetchNewbie = mysql_num_rows($rsFetchNewbie);
	
	
	
	if ($row_rsFetchNewbie['ConfirmByDate']==chr(0)) {
	
		print "You cannot confirm your account at the moment as you do not have access to this feature";
		exit;
	
	} else {
		
		if ($row_rsFetchNewbie['bConfirmed'] <> 0) {
			$showtable=false;
		} else {
			$showtable=true;
		}
		
		mysql_select_db($database_Epwebdev, $Epwebdev);
		$query_Recordset1 = sprintf("select uname, uid, rank, level from xoops_2004_users where pass = md5('%s') and uname = '%s'",$row_rsFetchNewbie['md5Pass'],$row_rsFetchNewbie['Username']);
		$Recordset1 = mysql_query($query_Recordset1, $Epwebdev) or die(mysql_error());
		$row_Recordset1 = mysql_fetch_assoc($Recordset1);
		$totalRows_Recordset1 = mysql_num_rows($Recordset1);
	
		if ($row_rsFetchNewbie['xoops_userID']==-1) {
		
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_rsFetchNewbie = sprintf("update sysops set xoops_userID = '%d' where RecID = '%d'",$row_Recordset1['uid'], $nSysopID);
			$rsFetchNewbie = mysql_query($query_rsFetchNewbie, $projectalpha) or die(mysql_error());
	
		}
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsStationary = "select HTML from stationary where StationaryCode = 'WELCOMENOTE_ONLINE'";
		$rsStationary = mysql_query($query_rsStationary, $projectalpha) or die(mysql_error());
		$row_rsStationary = mysql_fetch_assoc($rsStationary);
		$totalRows_rsStationary = mysql_num_rows($rsStationary);
	
		$adminMessage = str_replace("/Username/", $row_rsFetchNewbie['Username'], $row_rsStationary['HTML']);
		$adminMessage = str_replace("/Password/", $row_rsFetchNewbie['md5Pass'], $adminMessage);
		$adminMessage = str_replace("/Email/", $row_rsFetchNewbie['Email'], $adminMessage);
		$adminMessage = str_replace("/ConfirmationCode/", $row_rsFetchNewbie['ConfirmationCode'], $adminMessage);
		$adminMessage = str_replace("/ConfirmByDate/", $row_rsFetchNewbie['ConfirmByDate'], $adminMessage);
		$adminMessage = str_replace("/Firstname/", $row_rsFetchNewbie['Firstname'], $adminMessage);
		$adminMessage = str_replace("/Surname/", $row_rsFetchNewbie['Surname'], $adminMessage);
		$adminMessage = str_replace("/SysopID/", $nSysopID, $adminMessage);
		$adminMessage = str_replace("/SessionID/", session_id(), $adminMessage);
		

		$subject = sprintf("Activation Key for The Nexus 2005 - %s",$row_rsFetchNewbie['Username']);
		$xoopsMailer = & getMailer();
		$xoopsMailer->useMail();
		$xoopsMailer->setToEmails($row_rsFetchNewbie['Email']);
		$xoopsMailer->setFromEmail('admin@projectalpha.com.au');
		$xoopsMailer->setFromName('The Nexus');
		$xoopsMailer->setSubject($subject);
		$xoopsMailer->setBody($adminMessage);
		$xoopsMailer->send();
		if ( !$xoopsMailer->send() ) {
			echo $xoopsMailer->getErrors();
		}
		$showtable=true;
		
	}
} else {
	
	mysql_select_db($database_projectalpha, $projectalpha);
	$query_rsFetchNewbie = "select RecID, Username, md5(decode(`Password`,'dr34mt1me')) as md5Pass from sysops where ConfirmationCode = '$ConfirmationCode' and Email = '$Email' and RecID = '$nSysopID'";
	$rsFetchNewbie = mysql_query($query_rsFetchNewbie, $projectalpha) or die(mysql_error());
	$row_rsFetchNewbie = mysql_fetch_assoc($rsFetchNewbie);
	$totalRows_rsFetchNewbie = mysql_num_rows($rsFetchNewbie);

	if ($totalRows_rsFetchNewbie<>0) {
	 
		mysql_select_db($database_Epwebdev, $Epwebdev);
		$query_Recordset1 = sprintf("select uname, uid, rank, level from xoops_2004_users where pass = '%s' and uname = '%s'",$row_rsFetchNewbie['md5Pass'],$row_rsFetchNewbie['Username']);
		$Recordset1 = mysql_query($query_Recordset1, $Epwebdev) or die(mysql_error());
		$row_Recordset1 = mysql_fetch_assoc($Recordset1);
		$totalRows_Recordset1 = mysql_num_rows($Recordset1);
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsFetchNewbie = sprintf("update sysops set phpsessionid='', bConfirmed = '-1', ConfirmByDate = DATE_ADD(NOW(), INTERVAL 1 DAY) where RecID = '%d'",$row_rsFetchNewbie['RecID']);
		$rsFetchNewbie = mysql_query($query_rsFetchNewbie, $projectalpha) or die(mysql_error());
		
		
		$showtable=false;
	
	} else {
	
		echo "Perhaps you should not be tampering with this nice little server side file.";
		exit;
	}
	
	

}

?>





<body bgcolor="#FFFFFF" text="#000000">
<?php
include("top.php3");
?>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#66CC99">
  <tr>
    <td width="157" height="1307" align="center" valign="top" bgcolor="#FFFFFF"><img src="../images/idents/ep_ident.jpg" width="150" height="150"></td>
    <td width="613" valign="top" bgcolor="#FFCCCC">      <div align="center"><a href="login.php"><img src="../images/photos/img1.jpg" width="351" height="394" border="0"></a><br>
        Click Image to Logon<br>
        <span class="style24"><br>
      Welcome to the Nexus. This exciting new way of ebusiness and business resourcing will allow both us and generations to follow a new horizon in expanding business product and venture.</span> <br>
</div>
      <table width="90%"  border="0" align="center">
          <tr bgcolor="#000000">
            <td colspan="2"><div align="center"><span class="style24">The Nexus System Operator Account </span></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="33%"><span class="style6"><span class="style33">Sysop Identify Number:</span></span></td>
            <td width="67%"><div align="right" class="style13">
                <div align="center"><span class="style34"><?php echo $nSysopID; ?></span></div>
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style6"><span class="style33">Sysop Username: </span></span></td>
            <td><div align="right" class="style19 style35">
                <div align="center"><span class="style31"><?php echo $row_rsFetchNewbie['Username']; ?></span></div>
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style36">Sysop Description: </span></td>
            <td><div align="justify" class="style36">
                <div align="center" class="style22"><?php echo $row_rsFetchNewbie['Description']; ?></div>
            </div></td>
          </tr>
          <tr>
            <td colspan="2"><div align="right"><span class="style25"></span></div></td>
          </tr>
          <tr bgcolor="#000000" class="txtbox">
            <td colspan="2"><div align="center"><span class="style26">Nexus Business Center Forum Account </span></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style38">Forum Username:</span></td>
            <td><div align="center" class="style30 style19"><span class="style31"><?php echo $row_Recordset1['uname']; ?></span></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style38">Password:</span></td>
            <td><div align="center" class="style38"><em>same as sysop password registered now. </em></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style38">Level:</span></td>
            <td><div align="center" class="style38"><?php echo $row_Recordset1['level']; ?></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><span class="style38">Rank:</span></td>
            <td><div align="center" class="style38"><?php echo $row_Recordset1['rank']; ?></div></td>
          </tr>		 
      </table>
        <?php
	    if ($showtable==true) { ?>		<form name="form1" method="post" action="">
		  <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td colspan="2" bgcolor="#000000" class="style34"><div align="center" class="style41">email confirmation - Activation Code </div></td>
          </tr>
          <tr bgcolor="#000099">
            <td width="39%">&nbsp;</td>
            <td width="61%">&nbsp;</td>
          </tr>
          <tr bgcolor="#000099">
            <td class="style24"><div align="right">Email Activation Code: </div></td>
            <td>
              <div align="center">
                <input name="ConfirmationCode" type="text" class="txtbox" id="ConfirmationCode" size="40">
              </div></td>
          </tr>
          <tr bgcolor="#000099">
            <td class="txtbox"><div align="right"><span class="style43"><span class="style44">Email</span> Address:</span></div></td>
            <td><div align="center">
              <input name="Email" type="text" class="txtbox" id="Email" size="40">
            </div></td>
          </tr>
          <tr bgcolor="#000099">
            <td height="85"><div align="center"><span class="style48 style3">You have until <?php echo $row_rsFetchNewbie['ConfirmByDate']; ?> to activate your sysop account. This allows you to add your own company to the business network. </span></div></td>
            <td><div align="center">
              <p align="right"><span class="style45"><span class="style47">Report Email Address Abuse to admin@projectlpha.com.au<br> 
                </span><br>
                  <input type="submit" name="Submit" value="Go ahead activate me!">
              </span>                <br>
              </p>
              </div></td>
          </tr>
        </table>
		  <div align="center"><br>
		  Press 'Refresh' to resend your activation code.
          </div>
        </form><?php }  ?>
		</p></td>
  </tr>
</table>
<?php
	include('bottom.php3');
	?>
	
<p class="style1">&nbsp;</p>
<p class="style1">&nbsp;</p>
<p class="style1">&nbsp;</p>
<p class="style1">&nbsp;</p>

</body>
</html><?php
mysql_free_result($rsFetchNewbie);

mysql_free_result($rsStationary);
if (!empty($totalRows_Recordset1)) {
	mysql_free_result($Recordset1);
	}
?>
