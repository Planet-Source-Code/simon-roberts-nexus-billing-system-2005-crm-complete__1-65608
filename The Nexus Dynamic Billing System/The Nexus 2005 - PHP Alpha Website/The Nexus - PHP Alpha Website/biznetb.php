<?php require_once('connections/projectalpha.php'); ?>
<?php

$hn = gethostbyaddr($REMOTE_ADDR);
if (stristr($hn, "no1.com.au")) {
Header("HTTP/1.1 404 Not Found");
print "<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML 2.0//EN\">
	<HTML><HEAD>
	<TITLE>404 Not Found</TITLE>
	</HEAD><BODY>
	<H1>Not Found</H1>
	The requested URL / was not found on this server, Or you do not have access or your domain is banned..<P>
	<P>Additionally, a 404 Not Found
	error was encountered while trying to use an ErrorDocument to handle the request.
	<HR>
	<ADDRESS>Apache/1.3.26 Server at <your_domain> Port 80</ADDRESS>
	</BODY></HTML>";
	exit;
}

if(!session_id()){
  session_start();
}

mysql_select_db($database_projectalpha, $projectalpha);
$query_abv = "SELECT  count(left(virtualisp.Description,2)) as numocc, left(virtualisp.Description,2) as abvdesc FROM virtualisp group by left(virtualisp.Description,2) ORDER BY virtualisp.Description";
$abv = mysql_query($query_abv, $projectalpha) or die(mysql_error());
$totalRows_abv = mysql_num_rows($abv);

if ($nStart < 0) {
	$nStart = $nStart = 0;
}

if ($nNumRet <= 0) {
	$nNumRet = $nNumRet = 10;
}

mysql_select_db($database_projectalpha, $projectalpha);
$query_ViSPs = sprintf("SELECT * FROM virtualisp ORDER BY virtualisp.Description LIMIT %d, %d",$nStart, $nNumRet);
$ViSPs = mysql_query($query_ViSPs, $projectalpha) or die(mysql_error());
$totalRows_ViSPs = mysql_num_rows($ViSPs);
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Resellers on our Network</title>
<meta http-equiv="CACHE-CONTROL" content="PUBLIC; text/html; charset=iso-8859-1">
<META NAME="GOOGLEBOT" CONTENT="ALL">
<META NAME="ROBOTS" CONTENT="ALL"> 
<META NAME="KEYWORDS" 
CONTENT="Exitstencil Press Pty Ltd, Dolphin Communications, Dolphin Soltuions, ADSL, SHDSL, SH-DSL, gSHDSL, Broadband, Telstra, XOOPS, WOOT, NO1, Comcen, TPG, VOIP, SIP, PABX, Mainframe, AS/400, zSeries, DMT, MDMA, Amphetamine, Phone Systems, Simon Roberts, Jarrett Costi, Brad Bulters, Andrew Riddock, The Nexus, Mobile, Telco, Telecom, AAPT, Vodophone, Telstra, Hudsonson, Satellite, T1, Maths, Science, i-ching, philopshy<?php
	while($row_ViSPs = mysql_fetch_assoc($ViSPs)) {
		echo sprintf("%s, %s, %s, ",$row_ViSPs['Description'],$row_ViSPs['ABN'],$row_ViSPs['ACN']);
		}
		mysql_free_result($ViSPs);
		$ViSPs = mysql_query($query_ViSPs, $projectalpha) or die(mysql_error());
		?>">
<META HTTP-EQUIV="EXPIRES" 
CONTENT="31 Jul 2005 24:00:00 GMT 10+">
<META NAME="DESCRIPTION" 
CONTENT="The Nexus 2004 - Business Networks Currently In the Reseller & Wholesaler Interface.">
<META NAME="COPYRIGHT" CONTENT="&copy; 1972 - 2004 - Exitstencil Press Pty Ltd - All Rights Reserved.">
<META NAME="AUTHOR" CONTENT="Simon A. Roberts, Sydney, Australia - admin@projectalpha.com.au">
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #CCCCCC;
}
body {
	background-color: #<?php 
		if (empty($bckcolour)) {
			echo "990000";
		} else {
			echo sprintf("%s",$bckcolour);
		} ?>;
	margin-left: 020px;
	margin-top: 2px;
	margin-right: 20px;
	margin-bottom: 1px;
	background-image: url();
}
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
a:link {
	color: #00CCCC;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #00CCFF;
}
a:hover {
	text-decoration: none;
	color: #0099FF;
}
a:active {
	text-decoration: none;
	color: #00FFFF;
}
.style3 {font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
.style4 {color: #0099FF}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #0099FF; }
.style7 {font-size: 24}
.style14 {
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style20 {color: #FFFFFF}
.style22 {font-size: large}
.style23 {color: #FFFFFF; font-size: large; }
.style24 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; font-size: large; }
.style27 {
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
	font-weight: bold;
}
.style28 {
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
	color: #CCCCCC;
}
.style30 {color: #FFFFFF; font-size: large; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; }
.style34 {
	font-size: 10;
	font-weight: bold;
}
.style36 {font-size: 9px}
-->
</style></head>

<body>
<table width="95%"  border="0" align="center">
  <tr>
    <td colspan="2"><div align="justify">
      <p align="center" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="436" height="100">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="482" height="52">
          <param name="movie" value="biznet2.swf">
          <param name="quality" value="high">
          <param name="bgcolor" value="#<?php 
		if (empty($bckcolour)) {
			echo "990000";
			$bckcolour == "990000";
			$bckcolour = "990000";
		} else {
			echo sprintf("%s",$bckcolour);
		} ?>">
          <embed src="biznet2.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="482" height="52" bgcolor="#990000"></embed>
        </object>
        <span class="style30">Reseller utilising the amazing advantages of this software. ISP's, Teleco, Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP, Video Rental. </span></p>
    </div></td>
    <td colspan="5">
      <div align="center">
        <?php include(sprintf("%s%s","http://www.projectalpha.com.au/newsfeed.php?nVirtualID=0&level=10&bckcolour=",$bckcolour)); ?>
      </div></td></tr>
  <tr>
    <td width="4%"><strong><span class="style28"><br>
          </span></strong></td>
    <td colspan="5"><table width="100%" height="100%" border="0" align="center">
      <tr>
        <td width="88%" class="style28  style34"><div align="justify" class="style20"><span class="style36">Are you looking for a dynamic plateform, with remote daily backup of files and data. Barcode facility, fully networkable with any Unix/Win32 Hybrid platforms.
            
            Full 
            XML Embedded System - SSL Tunneling between you and the server. This including all Domain Documents &amp; HTML Produced Documents (Stationary/Reports) which can be ripped apon arival at any email destination or printed as a standard document, for record keeping... All safely store on a content lan in a major node in australia's CBD. 
          </span> </div></td>
        <td width="12%"><div align="center">
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="114" height="22">
            <param name="BGCOLOR" value="#<?php 
		if (empty($bckcolour)) {
			echo "990000";
		} else {
			echo sprintf("%s",$bckcolour);
		} ?>">
            <param name="movie" value="dc0001.swf">
            <param name="quality" value="high">
            <embed src="dc0001.swf" width="114" height="22" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#990000" ></embed>
          </object>
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="114" height="22">
            <param name="movie" value="bizforum.swf">
            <param name="quality" value="high">
            <param name="bgcolor" value="#<?php 
		if (empty($bckcolour)) {
			echo "#990000";
		} else {
			echo sprintf("%s",$bckcolour);
		} ?>">
            <embed src="bizforum.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="114" height="22" bgcolor="#990000"></embed>
          </object>
</div></td>
      </tr>
    </table></td>
    <td width="5%">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="7"  ><p>&nbsp;</p>    <table width="89%"  border="0" align="center" style="border-bottom: groove; border-left: groove; border-right: groove; border-top: groove">
      <tr bordercolor="#000000" bgcolor="#000000">
        <td width="17%"><div align="center" class="style1 style20 style22">
          <h5><strong>Emblem</strong></h5>
        </div></td>
        <td width="45%"> <div align="center" class="style1 style20 style22">
          <h5><strong>Description </strong></h5>
        </div></td>
        <td width="19%"><div align="center" class="style23">
          <h5><span class="style3">Number of Sysops </span></h5>
        </div></td>
        <td width="19%"><div align="center" class="style24">
          <h5><strong>Number of Products on Offer </strong></h5>
        </div></td>
        </tr>
	   <?php 
	  while ($row_ViSPs = mysql_fetch_assoc($ViSPs)) {
  ?>
      <tr bordercolor="#999999" bgcolor="#CCCCCC">
        <td rowspan="2" bordercolor="#CCCCCC"><div align="center" class="style1 style4 style4">
          <h4><a href="<?php echo sprintf("/bnpdossier.php?nVirtualID=%s", $row_ViSPs['RecID']); ?>"><?php  if (empty( $row_ViSPs['LogoURL'])) {
		  echo $row_ViSPs['ABN'];
      }else{
      echo sprintf('<img src="%s" width="160" height="160" border="0">', $row_ViSPs['LogoURL']);
      } ?></a></h4>
        </div>          <div align="left"></div></td>
        <td bordercolor="#CCCCCC"><div align="center" class="style1 style4 style7">          
          <h2><a href=<?php echo sprintf("/bnpdossier.php?nVirtualID=%s", $row_ViSPs['RecID']); ?> target="_self"><?php echo $row_ViSPs['Description']; ?></a></h2>
        </div></td>
        <td bordercolor="#CCCCCC"><div align="center" class="style5 style4">
          <h4><?php 
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_SysopCont = sprintf("SELECT count(*) as SysopsNo FROM sysops, virtualisp WHERE sysops.VirtualID = virtualisp.RecID and virtualisp.RecID = %s",$row_ViSPs['RecID']);
			$SysopCont = mysql_query($query_SysopCont, $projectalpha) or die(mysql_error());
			$row_SysopCont = mysql_fetch_assoc($SysopCont);
			$totalRows_SysopCont = mysql_num_rows($SysopCont);
			echo $row_SysopCont['SysopsNo']; ?></h4>
        </div></td>
		<?php
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_Earnings = sprintf("SELECT sum(invoiceout.AmountDue - invoiceout.AmountRefunded) as GrossErn, sum(invoiceout.GSTCharged - invoiceout.GSTRefunded) as GSTErn FROM invoiceout, virtualisp WHERE invoiceout.VirtualID = virtualisp.RecID and virtualisp.RecID = %s",$row_ViSPs['RecID']);
			$Earnings = mysql_query($query_Earnings, $projectalpha) or die(mysql_error());
			$row_Earnings = mysql_fetch_assoc($Earnings);
			$totalRows_Earnings = mysql_num_rows($Earnings);
		?>
        <td bordercolor="#CCCCCC"><div align="center" class="style5 style4">
          <h4><?php mysql_select_db($database_projectalpha, $projectalpha);
					$query_ProdNo = sprintf("SELECT count(*) as ProductNo FROM plantypes, virtualisp WHERE plantypes.VirtualID = virtualisp.RecID and virtualisp.RecID = %s",$row_ViSPs['RecID']);
					$ProdNo = mysql_query($query_ProdNo, $projectalpha) or die(mysql_error());
					$row_ProdNo = mysql_fetch_assoc($ProdNo);
					$totalRows_ProdNo = mysql_num_rows($ProdNo);
					echo $row_ProdNo['ProductNo']; ?></h4>
        </div></td>
        </tr>
      <tr bgcolor="#CCCCCC">
        <?php
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_Recordset4 = sprintf("SELECT  visp_phonenumbers.Extension, visp_phonenumbers.ContactName, visp_phonenumbers.PhoneNumber FROM visp_phonenumbers WHERE `visp_RecID` = '%d' and visp_phonenumbers.ContactName = 'sale'",$row_ViSPs['RecID']);
		$Recordset4 = mysql_query($query_Recordset4, $projectalpha) or die(mysql_error());
		$row_Recordset4 = mysql_fetch_assoc($Recordset4);
		$totalRows_Recordset4 = mysql_num_rows($Recordset4);
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_Recordset5 = sprintf("SELECT  visp_phonenumbers.Extension, visp_phonenumbers.ContactName, visp_phonenumbers.PhoneNumber FROM visp_phonenumbers WHERE `visp_RecID` = '%d' and visp_phonenumbers.ContactName = 'support'",$row_ViSPs['RecID']);
		$Recordset5 = mysql_query($query_Recordset5, $projectalpha) or die(mysql_error());
		$row_Recordset5 = mysql_fetch_assoc($Recordset5);
		$totalRows_Recordset5 = mysql_num_rows($Recordset5);

		mysql_select_db($database_projectalpha, $projectalpha);
		$query_Recordset6 = sprintf("SELECT  visp_phonenumbers.Extension, visp_phonenumbers.ContactName, visp_phonenumbers.PhoneNumber FROM visp_phonenumbers WHERE `visp_RecID` = '%d' and visp_phonenumbers.ContactName = 'account'",$row_ViSPs['RecID']);
		$Recordset6 = mysql_query($query_Recordset6, $projectalpha) or die(mysql_error());
		$row_Recordset6 = mysql_fetch_assoc($Recordset6);
		$totalRows_Recordset6 = mysql_num_rows($Recordset6);

		?>
        <td bordercolor="#999999" bgcolor="#999999" ><span class="style14 style20">Website: <a  href="http://www.<?php echo $row_ViSPs['Realm']; ?>">http://www.<?php echo $row_ViSPs['Realm']; ?>/</a>
		<br>Registered Business Number(s):<?php if (empty($row_ViSPs['ABN'])) {
		 } else { 
		 echo sprintf(' RBN: %s',$row_ViSPs['ABN']); } ?>
		 <?php
		 if (empty($row_ViSPs['ACN'])) { 
		 } else { 
		 echo sprintf(' ACN: %s',$row_ViSPs['ACN']); } ?>
		    <?php echo sprintf('<br>Sales: %s  ext. %s', $row_Recordset4['PhoneNumber'], $row_Recordset4['Extension']);  
		echo sprintf('<br>Support: %s  ext. %s', $row_Recordset5['PhoneNumber'], $row_Recordset5['Extension']); 
		echo sprintf('<br>Accounts: %s  ext. %s', $row_Recordset6['PhoneNumber'], $row_Recordset6['Extension']); 
		mysql_free_result($Recordset4);
		mysql_free_result($Recordset5);
		mysql_free_result($Recordset6);
		?>
		</span></td>
        <td bordercolor="#999999" bgcolor="#999999" >&nbsp;</td>
        <td bordercolor="#999999" bgcolor="#999999" >&nbsp;</td>
        </tr>
		  <?
	  }
	  ?>
    </table>
    <p>&nbsp;</p></td>
  </tr>
  <tr >
    <td><div align="right"></div></td>
    <td colspan="5" bgcolor="#666666"><div align="center"><span class="style27">
	   :[ Page ]:
        <?php 
		if ($nNumRet < 3) {
			$nNumRet=5;
		}
		$ki=0;
		while ($row_abv = mysql_fetch_assoc($abv)) {
			$indexchar = $row_abv['abvdesc'];
			$bi = 0;
			$abv_row = intval($abv_row) + 1;
			while (intval($bi) < intval($nNumRet) && !empty($row_abv)) {
				$abv_row = $abv_row + 1;
				$bi = intval($bi) + intval($row_abv['numocc']);
				$breaker = intval($row_abv['numocc']);
				$indexcharb = $row_abv['abvdesc'];
				$row_abv = mysql_fetch_assoc($abv);
			}
			if (intval($bi) > intval($nNumRet)) {
				echo sprintf("<a href=\"biznet.php?nStart=%d&nNumRet=%d\" target=\"_self\">", $ki, intval($nNumRet) + intval($breaker));  
				echo sprintf("[ %s - ",$indexchar);
				echo sprintf("%s ]</a> ",$indexcharb);
			} else {
				echo sprintf("<a href=\"biznet.php?nStart=%d&nNumRet=%d\" target=\"_self\">", $ki, intval($nNumRet));  
				echo sprintf("[ %s - ",$indexchar);
				echo sprintf("%s ]</a> ",$indexcharb);
				
			}
			$indexchar = $row_abv['abvdesc'];
			$ki = intval($ki) + intval($bi);
		}
		?>
    </span></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="35%">&nbsp;</td>
    <td width="53%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
<?php

mysql_free_result($ProdNo);

mysql_free_result($SysopCont);

mysql_free_result($Earnings);

mysql_free_result($abv);


mysql_free_result($ViSPs);

?>
