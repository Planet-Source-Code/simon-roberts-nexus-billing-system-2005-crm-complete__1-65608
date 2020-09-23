<?php require_once('Connections/projectalpha.php'); ?>
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
require_once('connections/projectalpha.php'); ?>
<?php

mysql_select_db($database_projectalpha, $projectalpha);
$query_ProductCount = "SELECT count(*) as Products FROM plantypes WHERE `VirtualID` = $nVirtualID and PeriodFee > 0";
$ProductCount = mysql_query($query_ProductCount, $projectalpha) or die(mysql_error());
$row_ProductCount = mysql_fetch_assoc($ProductCount);
$totalRows_ProductCount = mysql_num_rows($ProductCount);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset1 = "SELECT visp_emailaddresses.ContactName, visp_emailaddresses.EmailAddress FROM visp_emailaddresses WHERE `visp_RecID` = $nVirtualID limit 1,1";
$Recordset1 = mysql_query($query_Recordset1, $projectalpha) or die(mysql_error());
$row_Recordset1 = mysql_fetch_assoc($Recordset1);
$totalRows_Recordset1 = mysql_num_rows($Recordset1);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset2 = "SELECT visp_emailaddresses.ContactName, visp_emailaddresses.EmailAddress FROM visp_emailaddresses WHERE `visp_RecID` = $nVirtualID limit 2,1";
$Recordset2 = mysql_query($query_Recordset2, $projectalpha) or die(mysql_error());
$row_Recordset2 = mysql_fetch_assoc($Recordset2);
$totalRows_Recordset2 = mysql_num_rows($Recordset2);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset3 = "SELECT virtualisp.Realm FROM virtualisp WHERE `RecID` = $nVirtualID";
$Recordset3 = mysql_query($query_Recordset3, $projectalpha) or die(mysql_error());
$row_Recordset3 = mysql_fetch_assoc($Recordset3);
$totalRows_Recordset3 = mysql_num_rows($Recordset3);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset4 = "SELECT  visp_phonenumbers.Extension, visp_phonenumbers.ContactName, visp_phonenumbers.PhoneNumber FROM visp_phonenumbers WHERE `visp_RecID` = $nVirtualID limit 1,1";
$Recordset4 = mysql_query($query_Recordset4, $projectalpha) or die(mysql_error());
$row_Recordset4 = mysql_fetch_assoc($Recordset4);
$totalRows_Recordset4 = mysql_num_rows($Recordset4);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset5 = "SELECT visp_addresses.ContactName, visp_addresses.Street1, visp_addresses.Street2, visp_addresses.Country, visp_addresses.`State`, visp_addresses.Postcode, visp_addresses.Suburb, visp_addresses.PhotoURL FROM visp_addresses WHERE `visp_RecID` = $nVirtualID limit 1,1";
$Recordset5 = mysql_query($query_Recordset5, $projectalpha) or die(mysql_error());
$row_Recordset5 = mysql_fetch_assoc($Recordset5);
$totalRows_Recordset5 = mysql_num_rows($Recordset5);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset6 = "SELECT virtualisp.BriefDesc, virtualisp.LogoURL, virtualisp.Comment, virtualisp.CreatedBy_SysopID, virtualisp.VirtualID, virtualisp.SysopID, virtualisp.bTaxMode, virtualisp.cTaxCode, virtualisp.cTaxCountry, virtualisp.ABN, virtualisp.ACN, virtualisp.Description FROM virtualisp WHERE `RecID` = $nVirtualID";
$Recordset6 = mysql_query($query_Recordset6, $projectalpha) or die(mysql_error());
$row_Recordset6 = mysql_fetch_assoc($Recordset6);
$totalRows_Recordset6 = mysql_num_rows($Recordset6);

mysql_select_db($database_projectalpha, $projectalpha);
$query_CreatedBy = sprintf("SELECT sysops.Username, sysops.Firstname, sysops.Surname FROM sysops where RecID = %d", $row_Recordset6['CreatedBy_SysopID']);
$CreatedBy = mysql_query($query_CreatedBy, $projectalpha) or die(mysql_error());
$row_CreatedBy = mysql_fetch_assoc($CreatedBy);
$totalRows_CreatedBy = mysql_num_rows($CreatedBy);

mysql_select_db($database_projectalpha, $projectalpha);
$query_PrimarySysop = sprintf("SELECT sysops.Username, sysops.Firstname, sysops.Surname FROM sysops where RecID = %d", $row_Recordset6['SysopID']);
$PrimarySysop = mysql_query($query_PrimarySysop, $projectalpha) or die(mysql_error());
$row_PrimarySysop = mysql_fetch_assoc($PrimarySysop);
$totalRows_PrimarySysop = mysql_num_rows($PrimarySysop);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Sponsor = sprintf("SELECT RecID, Description FROM virtualisp WHERE RecID = %d",$row_Recordset6['VirtualID']);
$Sponsor = mysql_query($query_Sponsor, $projectalpha) or die(mysql_error());
$row_Sponsor = mysql_fetch_assoc($Sponsor);
$totalRows_Sponsor = mysql_num_rows($Sponsor);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset9 = "SELECT distinct AVG((plantypes.PeriodFee * tax.Percentage) + plantypes.PeriodFee) as AVGCycle, STDDEV((plantypes.PeriodFee * tax.Percentage) + plantypes.PeriodFee) as STDCycle, AVG(plantypes.MBPerPeriod) as AVGMBPrice, AVG(plantypes.MBBlockSize) as AVGMBSize, AVG(plantypes.FeePerBlock) as AVGMBBlock, AVG((plantypes.JoiningFee * tax.Percentage) + plantypes.JoiningFee)as AVGJoin, STDDEV((plantypes.JoiningFee * tax.Percentage) + plantypes.JoiningFee) as STDJoin FROM plantypes inner join virtualisp On virtualisp.RecID = plantypes.VirtualID inner join tax on virtualisp.cTaxCode = tax.Code and virtualisp.cTaxCountry = tax.Country WHERE plantypes.VirtualID = '$nVirtualID' and plantypes.PeriodFee > 0";
$Recordset9 = mysql_query($query_Recordset9, $projectalpha) or die(mysql_error());
$row_Recordset9 = mysql_fetch_assoc($Recordset9);
$totalRows_Recordset9 = mysql_num_rows($Recordset9);


mysql_select_db($database_projectalpha, $projectalpha);
$query_topseller = "SELECT distinct count(acci_services.RecID) as ttlcnt, plantypes.RecID, plantypes.CatNo, plantypes.Description, plantypes.BillOnce, plantypes.PeriodFee, (plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee as TaxInc,  (((plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee) /plantypes.chgIntervalDays) as DailyTaxInc,plantypes.chgPerTxt,(plantypes.JoiningFee* tax.Percentage)+plantypes.JoiningFee as JoiningInc, plantypes.JoiningFee, servicetypes.Description as SVrDESC FROM plantypes inner join virtualisp ON plantypes.VirtualID = virtualisp.RecID inner join tax inner join servicetypes on plantypes.ServiceID = servicetypes.RecID, acci_services WHERE acci_services.ptRecID = plantypes.RecID and virtualisp.cTaxCode = tax.Code and virtualisp.cTaxCountry = tax.Country and plantypes.PeriodFee > 0 and acci_services.VirtualID = '$nVirtualID' GROUP BY plantypes.RecID ORDER BY plantypes.PeriodFee ASC";
$topseller = mysql_query($query_topseller, $projectalpha) or die(mysql_error());
$row_topseller = mysql_fetch_assoc($topseller);
$totalRows_topseller = mysql_num_rows($topseller);
$btopseller = true;

if ($totalRows_topseller == 0) {
	$btopseller = false;
	} else {
			while ($row_topseller = mysql_fetch_assoc($topseller)) {
			if ($row_topseller['ttlcnt'] >= $ttlcnt) {
				$ttlcnt = $row_topseller['ttlcnt'];
				$query_topseller = sprintf("SELECT distinct plantypes.RecID, plantypes.CatNo, plantypes.Description, plantypes.BillOnce, plantypes.PeriodFee, (plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee as TaxInc,  (((plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee) /plantypes.chgIntervalDays) as DailyTaxInc,plantypes.chgPerTxt,(plantypes.JoiningFee* tax.Percentage)+plantypes.JoiningFee as JoiningInc, plantypes.JoiningFee, servicetypes.Description as SVrDESC FROM plantypes inner join virtualisp ON plantypes.VirtualID = virtualisp.RecID inner join tax inner join servicetypes on plantypes.ServiceID = servicetypes.RecID, acci_services WHERE virtualisp.cTaxCode = tax.Code and virtualisp.cTaxCountry = tax.Country and plantypes.PeriodFee > 0 and acci_services.VirtualID = '%s' and plantypes.RecID ='%s' GROUP BY plantypes.RecID ORDER BY plantypes.ServiceID",$nVirtualID,$row_topseller['RecID']);
				
			}
			
		}
	}

mysql_free_result($topseller);

if ($btopseller == true) {
	mysql_select_db($database_projectalpha, $projectalpha);
	//print $query_topseller;
	$topseller = mysql_query($query_topseller, $projectalpha) or die(mysql_error());
	$row_topseller = mysql_fetch_assoc($topseller);
	$totalRows_topseller = mysql_num_rows($topseller);
}

	
mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset7 = "SELECT distinct plantypes.RecID, plantypes.CatNo, plantypes.Description, plantypes.BillOnce, plantypes.PeriodFee, (plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee as TaxInc,  (((plantypes.PeriodFee* tax.Percentage)+plantypes.PeriodFee) /plantypes.chgIntervalDays) as DailyTaxInc,plantypes.chgPerTxt,(plantypes.JoiningFee* tax.Percentage)+plantypes.JoiningFee as JoiningInc, plantypes.JoiningFee, servicetypes.Description as SVrDESC FROM plantypes inner join virtualisp ON plantypes.VirtualID = virtualisp.RecID inner join tax inner join servicetypes on plantypes.ServiceID = servicetypes.RecID WHERE virtualisp.cTaxCode = tax.Code and virtualisp.cTaxCountry = tax.Country and plantypes.PeriodFee > 0 and plantypes.VirtualID = '$nVirtualID' ORDER BY plantypes.ServiceID";
$Recordset7 = mysql_query($query_Recordset7, $projectalpha) or die(mysql_error());
$totalRows_Recordset7 = mysql_num_rows($Recordset7);

mysql_select_db($database_projectalpha, $projectalpha);
$query_Recordset8 = "SELECT distinct plantypes.RecID FROM plantypes inner join virtualisp ON plantypes.VirtualID = virtualisp.RecID inner join tax inner join servicetypes on plantypes.ServiceID = servicetypes.RecID WHERE virtualisp.cTaxCode = tax.Code and virtualisp.cTaxCountry = tax.Country and plantypes.PeriodFee > 0 and plantypes.VirtualID = '$nVirtualID' ORDER BY plantypes.ServiceID";
$Recordset8 = mysql_query($query_Recordset8, $projectalpha) or die(mysql_error());
$totalRows_Recordset8 = mysql_num_rows($Recordset7);

?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><?php echo $row_Recordset6['Description']; ?> - <?php echo $row_Recordset6['ABN']; ?></title>
<meta http-equiv="CACHE-CONTROL" content="PUBLIC; text/html; charset=iso-8859-1">
<META NAME="GOOGLEBOT" CONTENT="ALL">
<META NAME="ROBOTS" CONTENT="ALL"> 
<META NAME="KEYWORDS" 
CONTENT="<?php echo $row_Recordset6['Description']; ?> - <?php echo $row_Recordset6['ABN']; ?>,Exitstencil Press Pty Ltd, Dolphin Communications, Dolphin Soltuions, ADSL, SHDSL, SH-DSL, gSHDSL, Broadband, Telstra, XOOPS, WOOT, NO1, Comcen, TPG, VOIP, SIP, PABX, Mainframe, AS/400, zSeries, DMT, MDMA, Amphetamine, Phone Systems, Simon Roberts, Jarrett Costi, Brad Bulters, Andrew Riddock, The Nexus, Mobile, Telco, Telecom, AAPT, Vodophone, Telstra, Hudsonson, Satellite, T1, Maths, Science, i-ching, philopshy">
<META HTTP-EQUIV="EXPIRES" 
CONTENT="31 Jul 2005 24:00:00 GMT 10+">
<META NAME="DESCRIPTION" 
CONTENT="The Nexus 2005 - <?php echo $row_Recordset6['Description']; ?> - <?php echo $row_Recordset6['ABN']; ?> dossier.">
<META NAME="COPYRIGHT" CONTENT="&copy; 1972 - 2004 - Exitstencil Press Pty Ltd - All Rights Reserved.">
<META NAME="AUTHOR" CONTENT="Simon A. Roberts, Sydney, Australia - admin@projectalpha.com.au">
<style type="text/css">
<!--
body {
	background-image:  url();
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	background-color: #8D5061;
	margin-bottom: 5px;
}
.FontMode1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 24px;
}
.FontMode2 {font-family: Arial, Helvetica, sans-serif}
.FontMode3 {font-size: 24px}
.FontMode4 {
	color: #009999;
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
}
.FontMode5 {
	color: #000066;
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
}
.FontMode9 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #336600;
}
.FontMode11 {color: #FFFFFF}
.FontMode26 {font-family: Arial, Helvetica, sans-serif; color: #FFFFFF; font-weight: bold; font-size: medium; }
.FontMode36 {
	color: #000000;
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
}
.FontMode38 {font-size: medium; font-weight: bold; }
.FontMode40 {font-size: medium; font-weight: bold; color: #FFFFFF; }
.FontMode44 {color: #FFFFFF; font-size: 9px; font-family: Arial, Helvetica, sans-serif; }
.FontMode48 {color: #FFFFFF; font-size: 10px; font-family: Arial, Helvetica, sans-serif; }
.FontMode54 {font-family: Geneva, Arial, Helvetica, sans-serif}
.FontMode56 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"}
.FontMode57 {font-size: 12px}
.FontMode62 {font-size: 24px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; }
.FontMode64 {font-size: 24px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-weight: bold; }
.FontMode65 {color: #000000}
.FontMode66 {font-weight: bold; font-size: 45px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";}
h1,h2,h3,h4,h5,h6 {
	font-family: Tahoma, Trebuchet MS, Verdana, Lucida Console;
}
.FontMode75 {font-size: medium; font-weight: bold; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; }
.FontMode76 {
	color: #990000;
	font-size: 12px;
}
.FontMode78 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-size: 12px; color: #FF9933; }
.FontMode79 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-size: 12px; font-weight: bold; }
body,td,th {
	font-family: Tahoma, Trebuchet MS, Verdana, Lucida Console;
	font-size: 14px;
}
.style2 {color: #66CCFF}
.style3 {font-size: 12px; color: #66CCFF; }
.style4 {
	font-size: 14px;
	color: #CCCCFF;
}
.style6 {
	color: #FF9933;
	font-weight: bold;
}
.style7 {
	color: #FFCC33;
	font-weight: bold;
}
.style10 {color: #FF9933}
.style13 {
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
	font-size: xx-large;
	color: #FF9933;
	font-weight: bold;
}
-->
</style><meta http-equiv="Content-Type" content="text/html; charset="></head>

<body>
<table width="94%" height="185"  border="0" cellpadding="0" cellspacing="1" align="center">
      <tr bgcolor="#000000" class="FontMode1">
        <td width="66%" height="29" ><div align="right" class="FontMode65">
          <div align="center"><span class="FontMode2">        <span class="FontMode54"><span class="FontMode66">
            <?php  if (empty( $row_Recordset6['LogoURL'])) {
		  print 'No logo Set';
      }else{
      echo sprintf('<img src="%s" width="250" height="250" align="top" >', $row_Recordset6['LogoURL']);
      } ?>
          </span></span></span></div>
        </div></td>
          <td width="34%" height="29"><div align="center" class="FontMode65 FontMode11">
            <p class="FontMode2">                  <span class="style13"><?php echo $row_Recordset6['Description']; ?> </span><span class="FontMode56"><span class="FontMode75"><span class="FontMode57"><span class="style4"><br />
                    </span>        <span class="style2"><br />
            Registration Details </span></span></span><span class="style3"><br />
                  <strong><strong>ABN/RBN: <?php echo $row_Recordset6['ABN']; ?><br />
                  <strong>ACN: <?php echo $row_Recordset6['ACN']; ?></strong></strong></strong></span></span> </p>
            <p class="style6"><span class="style6">
              <?php if ($btopseller == TRUE) { 
				echo "<p align=\"center\" >Top Selling Service from $";
				echo sprintf("%01.2f",$row_topseller['DailyTaxInc']); 
				echo $row_topseller['chgPerTxt']; 
				echo "<br />";
				echo $row_topseller['Description']; 
				echo "<br />Service Type:";
				echo $row_topseller['SVrDESC']; 
				echo "<br />Product Code: ";
				echo $row_topseller['CatNo']; 

				echo "<br />Cycle Fee: $ ";
				echo sprintf("%01.2f",$row_topseller['TaxInc']);
				echo ", ";
				if ($row_topseller['JoiningInc'] > 0) {
					echo "Joining Fee: $ ";
					echo sprintf("%01.2f",$row_topseller['JoiningInc']); 
				} else {

					echo "-[ Joining Free ]-";
				}					
				echo "<br /><br />"; } ?>
            </span> </span> </span> </p>
            </p>
            <span class="FontMode76 FontMode56  FontMode57 style6"><span class="FontMode76 FontMode56 FontMode57  style10">Sponsor Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP: </span></span><span class="FontMode76 FontMode56  FontMode57"><span class="style7"><a  href="http://www.projectalpha.com.au/vispdossier.php?nVirtualID=<?php echo $row_Sponsor['RecID']; ?>"><?php echo $row_Sponsor['Description']; ?></a></span></span></div></td>
      </tr>
  <tr class="FontMode1">
      <td height="29" colspan="2"><h2 align="right" class="FontMode9"><span class="FontMode26">       </span></h2>      </td>
  </tr>
  <tr class="FontMode1">
    <td height="27" colspan="2"><div align="center"><span class="FontMode2"><span class="FontMode44">
    </span></span></div></td>
  </tr>
  <tr class="FontMode1" bgcolor="#999999">
    <td height="27" colspan="2"><span class="FontMode40">Subscriptions Retailers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP </span></td>
  </tr>
  <tr class="FontMode1" bgcolor="#FFFFFF" >
    <td height="39" colspan="2">
      <div align="center" class="FontMode56">
        <table width="100%" border="0">
          <tr>
            <td width="634"><div align="center">
              <applet code="AcuteShifter.class" codebase="java/" name="labPanelApplet"
           width="100%" height="117" align="absmiddle" archive="AcuteShifter.jar" id="labPanelApplet">
                  <!-- The message that will be displayed in the applet. -->
                  <param name="Message" value="
			   <?php 
	  while ($row_Recordset7 = mysql_fetch_assoc($Recordset7)) {
  ?>

 		<?php
		 echo '<PROD';
		 echo $row_Recordset7['RecID']; 
		 echo '><h>';
		 echo $row_Recordset7['SVrDESC'];
		 if ($row_Recordset7['BillOnce'] == 1) {
				echo ' - billed only once - outright fee inc tax $ ';
	  	 	    echo sprintf("%01.2f",$row_Recordset7['PeriodFee']);
				echo '</h>';
			} else {
				echo ' - from as low as $ ';
				echo sprintf("%01.2f",$row_Recordset7['DailyTaxInc']);
			 	echo sprintf(" %s.</h>",$row_Recordset7['chgPerTxt']);
		 }
		 echo ' Product Code: ';
         echo $row_Recordset7['CatNo']; 
         echo '<br>';
		 echo ' Description: ';
		 echo $row_Recordset7['Description']; 
		 echo '<br><fullstory>';
 		 echo 'Excluding Tax: Subscription per cycle $ ';
		 echo sprintf("%01.2f",$row_Recordset7['PeriodFee']);
		 if ($row_Recordset7['JoiningFee'] > 0) {
			 echo ', Join/Setup Fee $ ';
 			 echo sprintf("%01.2f",$row_Recordset7['JoiningFee']);
		 }

         echo '<br>';
  		 echo 'Including Tax: Subscription per cycle $ ';
		 echo sprintf("%01.2f",$row_Recordset7['TaxInc']);
		 if ($row_Recordset7['JoiningInc'] > 0) {
			 echo ', Join/Setup Fee $ ';
			 echo sprintf("%01.2f",$row_Recordset7['JoiningInc']);
		 }
		 echo '</fullstory>';
		 echo '</PROD';
		 echo $row_Recordset7['RecID']; 
		 echo '>';
		 
			?>
<?php } ?>
		 

      " />
                  <!-- The definition of all FontModes used by the message. -->
                  <param name="style" value="
         <default
            Text-Size='12'
            Text-Color='0033CC'
            Shift-Pause='9000'
            Shift-In-Effect='slide-up'
            Padding-Top='5'
            Background-Image-Repeat='true'
            Background-Color='ffffff'
            Border-Color='ffffff'
            Border-Type='full'>
         <h Text-Size='12'
            Text-Color='84a5d5'
            Padding-Top='6'>
         </h
            Padding-Bottom='3'>
         <fullstory
            Text-Color='fb2200'
            Text-Color-Over='ffffff'
            Text-Color-Click='ffffff'
            Padding-Top='3'
            Line-Break='true'
            Align='right'>
			
<?php
			
             $sevenbit = $_SESSION['sevenbit'];   
	  while ($row_Recordset8 = mysql_fetch_assoc($Recordset8)) {
  ?>
		<?php
			
			++$sevenbit;
			if ($sevenbit == 1) {
				 echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-up\'>',$row_Recordset8['RecID']);
				 } else {
				 if ($sevenbit == 2) {
					echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-down\'>',$row_Recordset8['RecID']);
				 } else { 
					if ($sevenbit == 3) {
						echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-left\'>',$row_Recordset8['RecID']);
					
					} else {
						if ($sevenbit == 4) {
						 echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-right\'>',$row_Recordset8['RecID']);
						 } else {
						 if ($sevenbit == 5) {
							echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-down\'>',$row_Recordset8['RecID']);
						 } else { 
							if ($sevenbit == 6) {
								echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-up\'>',$row_Recordset8['RecID']);
							
							} else {
								echo sprintf('</PROD%s Section-Header=\'true\' Shift-In-Effect=\'slide-left\'>',$row_Recordset8['RecID']);
								while ($sevenbit > 0) {
									--$sevenbit;
										}
									}
								}
						   }
						
						}
					}
				}
		   
			 
			?>
	<?php }
		$_SESSION['sevenbit'] = $sevenbit; ?>

      " />
                  <!-- The following parameters are used to format the applet
           area while images and input files are loaded. (Optional).-->
                  <param name="Loading-Text" value="Loading plans currently on offer..." />
                  <param name="Loading-Text-Color" value="333333" />
                  <param name="Loading-Background-Color" value="f8f8f8" />
                  <!-- When you register AcuteApplets you will get Domain-Keys 
           that removes the intro nag-screen. -->
                  <PARAM name="Domain-Keys" value="13280,13213"/>
              </applet>
            </div></td>
            <td width="318"><p class="FontMode57">Average Subscription Cycle Fee: <?php echo sprintf("$ %01.2f",$row_Recordset9['AVGCycle']); ?><br />
              Average Joining Fee: <?php echo sprintf("$ %01.2f",$row_Recordset9['AVGJoin']); ?><br />
              Average Size of a MB Block: <?php echo sprintf("%01.4f MB's",$row_Recordset9['AVGMBSize']); ?></p>
              <p class="FontMode57">Standard Deviation of Cycle Fees: <?php echo sprintf("$ %01.2f",$row_Recordset9['STDCycle']); ?><br />
              Standard Deviation of Joining Fees: <?php echo sprintf("$ %01.2f",$row_Recordset9['STDJoin']); ?></p></td>
          </tr>
        </table>
</div></td>
  </tr>
  <tr class="FontMode1">
    <td height="27" colspan="2"><blockquote>
      <blockquote>
        <blockquote>
          <blockquote>
            <p>&nbsp;                            </p>
          </blockquote>
        </blockquote>
      </blockquote>
    </blockquote>      </td>
  </tr>
</table>
<table width="94%" height="464"  border="0" cellpadding="0" cellspacing="1" align="center">
  <tr class="FontMode2" bgcolor="#999999">
    <td height="20"><span class="FontMode38 FontMode11">Website:</span></td>
    <td><span class="FontMode40">Primary Email Address: </span></td>
    <td colspan="3"><span class="FontMode40">Secondary Email Address: </span></td>
  </tr>
  <tr class="FontMode2" bgcolor="#CCCCCC">
    <td width="26%" height="61"><div align="center" class="FontMode62 FontMode3"><strong><a href="http://www.<?php echo $row_Recordset3['Realm']; ?>">www.<?php echo $row_Recordset3['Realm']; ?></a></strong></div></td>
    <td width="37%"><p align="center" class="FontMode62"><strong><?php echo $row_Recordset1['ContactName']; ?></strong></p>
    <p align="center" class="FontMode64"><a href="mailto:<?php echo $row_Recordset1['EmailAddress']; ?>"><?php echo $row_Recordset1['EmailAddress']; ?></a></p></td>
    <td colspan="3"><p align="center" class="FontMode62"><strong><?php echo $row_Recordset2['ContactName']; ?></strong></p>
    <p align="center" class="FontMode64"><a href="mailto:<?php echo $row_Recordset2['EmailAddress']; ?>"><?php echo $row_Recordset2['EmailAddress']; ?></a></p></td>
  </tr>
  <tr class="FontMode2">
    <td height="27" class="FontMode3">&nbsp;</td>
    <td colspan="4">&nbsp;</td>
  </tr>
  <tr class="FontMode2" bgcolor="#999999">
    <td height="27" class="FontMode3"><span class="FontMode40">Sales Contact</span></td>
    <td colspan="4"><span class="FontMode40">Principle Place of Business </span></td>
  </tr>
  <tr class="FontMode2" bgcolor="#CCCCCC">
    <td height="199" class="FontMode3"><div align="center" class="FontMode5"><?php echo $row_Recordset4['ContactName']; ?><br />
      <?php echo $row_Recordset4['PhoneNumber']; ?><br />
    ( Ext. <?php echo $row_Recordset4['Extension']; ?> ) </div></td>
    <td colspan="4"><p align="center"><span class="FontMode3"></span><span class="FontMode3"></span><span class="FontMode3"></span><span class="FontMode3">	  <span class="FontMode56">
      <?php  if (empty( $row_Recordset5['PhotoURL'])) {
		  print 'No photo of business frontage taken...';
      }else{
      echo sprintf('<img src="%s" width="542" height="246" >', $row_Recordset5['PhotoURL']);
      } ?>
    </span></span></p>
      <p align="center" class="FontMode1 FontMode4"><?php echo $row_Recordset5['ContactName']; ?><br />
      <?php echo $row_Recordset5['Street1']; ?> <br />
      <?php echo $row_Recordset5['Street2']; ?> <br />
      <?php echo $row_Recordset5['Suburb']; ?>  <br />
      <?php echo $row_Recordset5['State']; ?>, <?php echo $row_Recordset5['Country']; ?>,,<?php echo $row_Recordset5['Postcode']; ?></p>      </td>
  </tr>
  <tr class="FontMode2">
    <td colspan="5">&nbsp;</td>
  </tr>
  <tr class="FontMode2">
    <td colspan="5"><div align="center"><span class="FontMode48">
        <?php if (empty($row_Recordset6['BriefDesc'])) {
	  	echo '<br>Company Description or RAW HTML not entered yet.<br>See Virtual ISP Configuration form.<br>';
	  } else {
			$pos      = strpos($row_Recordset6['BriefDesc'], 'http://');

			if ($pos === false) {
				echo $row_Recordset6['BriefDesc'];			   
			} else {
				if ($pos <= 2) {
					echo sprintf('<br>***  this is a live link to %s ****<br>',$row_Recordset6['BriefDesc']);
					include($row_Recordset6['BriefDesc']);
				} else {
					echo $row_Recordset6['BriefDesc'];			   
				}
		} } ?>
    </span></div></td>
  </tr>
  <tr class="FontMode2">
    <td colspan="5"><span class="FontMode3"></span><span class="FontMode3"></span><span class="FontMode3"></span><span class="FontMode3"></span></td>
  </tr>
  <tr class="FontMode2" bgcolor="#FFFFFF">
    <td height="95"><div align="center"><span class="FontMode3"><img src="images/icons/genpa.gif" width="136" height="36" /><br />
    </span></div></td>
    <td colspan="4"><span class="FontMode3"></span><span class="FontMode3"></span><span class="FontMode3"></span>      <ul class="FontMode36">
      <li><strong>This Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP was created by <?php echo $row_CreatedBy['Firstname']; ?> <?php echo $row_CreatedBy['Surname']; ?></strong></li>
      <li><strong>          They have a total of <?php echo $row_ProductCount['Products']; ?> products on offer.</strong></li>
      <li><strong> The primary administrator for this Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP&reg; is <?php echo $row_PrimarySysop['Firstname']; ?> <?php echo $row_PrimarySysop['Surname']; ?> </strong></li>
      <li><strong>Tax Code [ <?php echo $row_Recordset6['cTaxCode']; ?> ], Tax Country [ <?php echo $row_Recordset6['cTaxCountry']; ?> ] </strong></li>
    </ul></td>
  </tr>
</table>
</body>
</html>
<?php
mysql_free_result($CreatedBy);

mysql_free_result($ProductCount);

mysql_free_result($PrimarySysop);

mysql_free_result($Sponsor);

mysql_free_result($Recordset7);
mysql_free_result($Recordset8);
mysql_free_result($Recordset9);

mysql_free_result($topseller);

mysql_free_result($Recordset1);

mysql_free_result($Recordset2);

mysql_free_result($Recordset3);

mysql_free_result($Recordset4);

mysql_free_result($Recordset5);

mysql_free_result($Recordset6);

?>
