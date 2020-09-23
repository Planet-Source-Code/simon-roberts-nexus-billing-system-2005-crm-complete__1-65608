<?php require_once('../../Connections/projectalpha.php'); ?>
<?php

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPb = sprintf("SELECT accountinfo.SysopID, accountinfo.RecID FROM accountinfo inner join invoicetraxr on accountinfo.RecID = invoicetraxr.acci_RecID WHERE MD5(gPassword) = '%s' and invoicetraxr.RecID = %d",$gPass,$nTraxrID);
$rsVISPb = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
$row_rsVISPb = mysql_fetch_assoc($rsVISPb);
$totalRows_rsVISPb = mysql_num_rows($rsVISPb);

if ($totalRows_rsVISPb == 0) {
	echo 'You do not have access to this report. Please revise the link you have selected.';
	exit;
}
//echo "2<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPa = sprintf("SELECT * FROM virtualisp WHERE virtualisp.RecID = %d",$nVirtualID);
$rsVISPa = mysql_query($query_rsVISPa, $projectalpha) or die(mysql_error());
$row_rsVISPa = mysql_fetch_assoc($rsVISPa);
$totalRows_rsVISPa = mysql_num_rows($rsVISPa);

if (empty($row_rsVISPa['ACN'])) {
	if (empty($row_rsVISPa['ABN'])) {
		$RBN = 'No RBN Currently Set';
		} else {
			$RBN = "ABN " + $row_rsVISPa['ABN'];
		}
	} else {
			$RBN = "ACN " + $row_rsVISPa['ABN'];
		}

//echo "3<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsSysNow = "SELECT distinct NOW() as SysNow FROM sysops";
$rsSysNow = mysql_query($query_rsSysNow, $projectalpha) or die(mysql_error());
$row_rsSysNow = mysql_fetch_assoc($rsSysNow);
$totalRows_rsSysNow = mysql_num_rows($rsSysNow);

//echo "4<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsSponsor = sprintf("SELECT virtualisp.Description FROM virtualisp where RecID = '%d'",$row_rsVISPa['VirtualID']);
$rsSponsor = mysql_query($query_rsSponsor, $projectalpha) or die(mysql_error());
$row_rsSponsor = mysql_fetch_assoc($rsSponsor);
$totalRows_rsSponsor = mysql_num_rows($rsSponsor);

//echo "5<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsSysop = sprintf("SELECT sysops.SecurityLevel, sysops.Firstname, sysops.Surname, sysops.Email FROM sysops where RecID = '%d'",$row_rsVISPb['SysopID']);
$rsSysop = mysql_query($query_rsSysop, $projectalpha) or die(mysql_error());
$row_rsSysop = mysql_fetch_assoc($rsSysop);
$totalRows_rsSysop = mysql_num_rows($rsSysop);

//echo "6<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPAddyA = sprintf("SELECT visp_addresses.ContactName, visp_addresses.Street1, visp_addresses.Street2, visp_addresses.Country, visp_addresses.`State`, visp_addresses.Postcode, visp_addresses.Suburb, visp_addresses.visp_RecID FROM visp_addresses WHERE visp_addresses.ContactName like '%s' and visp_RecID = '%d' limit 1,1",'%sale%',$nVirtualID);
$rsVISPAddyA = mysql_query($query_rsVISPAddyA, $projectalpha) or die(mysql_error());
$row_rsVISPAddyA = mysql_fetch_assoc($rsVISPAddyA);
$totalRows_rsVISPAddyA = mysql_num_rows($rsVISPAddyA);

//echo "7<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPAddyb = sprintf("SELECT visp_addresses.ContactName, visp_addresses.Street1, visp_addresses.Street2, visp_addresses.Country, visp_addresses.`State`, visp_addresses.Postcode, visp_addresses.Suburb, visp_addresses.visp_RecID FROM visp_addresses where visp_addresses.ContactName like '%s' and visp_RecID = '%d' limit 1,1",'%account%',$nVirtualID);
$rsVISPAddyb = mysql_query($query_rsVISPAddyb, $projectalpha) or die(mysql_error());
$row_rsVISPAddyb = mysql_fetch_assoc($rsVISPAddyb);
$totalRows_rsVISPAddyb = mysql_num_rows($rsVISPAddyb);

//echo "8<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPAddyC = sprintf("SELECT visp_addresses.ContactName, visp_addresses.Street1, visp_addresses.Street2, visp_addresses.Country, visp_addresses.`State`, visp_addresses.Postcode, visp_addresses.Suburb, visp_addresses.visp_RecID FROM visp_addresses where visp_addresses.ContactName like '%s' and visp_RecID = '%d' limit 1,1",'%support%',$nVirtualID);
$rsVISPAddyC = mysql_query($query_rsVISPAddyC, $projectalpha) or die(mysql_error());
$row_rsVISPAddyC = mysql_fetch_assoc($rsVISPAddyC);
$totalRows_rsVISPAddyC = mysql_num_rows($rsVISPAddyC);

//echo "9<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPPhnA = sprintf("SELECT visp_phonenumbers.visp_RecID, visp_phonenumbers.DateAdded, visp_phonenumbers.PhoneNumber, visp_phonenumbers.Extension, visp_phonenumbers.ContactName FROM visp_phonenumbers where ContactName Like '%s' and visp_RecID = '%d'",'%sale%',$nVirtualID);
$rsVISPPhnA = mysql_query($query_rsVISPPhnA, $projectalpha) or die(mysql_error());
$row_rsVISPPhnA = mysql_fetch_assoc($rsVISPPhnA);
$totalRows_rsVISPPhnA = mysql_num_rows($rsVISPPhnA);

//echo "10<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPPhnB = sprintf("SELECT visp_phonenumbers.visp_RecID, visp_phonenumbers.DateAdded, visp_phonenumbers.PhoneNumber, visp_phonenumbers.Extension, visp_phonenumbers.ContactName FROM visp_phonenumbers where ContactName Like '%s' and visp_RecID = %d",'%account%',$nVirtualID);
$rsVISPPhnB = mysql_query($query_rsVISPPhnB, $projectalpha) or die(mysql_error());
$row_rsVISPPhnB = mysql_fetch_assoc($rsVISPPhnB);
$totalRows_rsVISPPhnB = mysql_num_rows($rsVISPPhnB);

//echo "11<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPPhnC = sprintf("SELECT visp_phonenumbers.visp_RecID, visp_phonenumbers.DateAdded, visp_phonenumbers.PhoneNumber, visp_phonenumbers.Extension, visp_phonenumbers.ContactName FROM visp_phonenumbers where ContactName Like '%s' and visp_RecID = %d",'%support%',$nVirtualID);
$rsVISPPhnC = mysql_query($query_rsVISPPhnC, $projectalpha) or die(mysql_error());
$row_rsVISPPhnC = mysql_fetch_assoc($rsVISPPhnC);
$totalRows_rsVISPPhnC = mysql_num_rows($rsVISPPhnC);

//echo "12<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsVISPPrimartAddy = sprintf("Select visp_addresses.Street1, visp_addresses.Street2, visp_addresses.Country, visp_addresses.`State`, visp_addresses.Postcode, visp_addresses.Suburb FROM visp_addresses WHERE visp_addresses.visp_RecID = '%d' limit 1,1",$nVirtualID);
$rsVISPPrimartAddy = mysql_query($query_rsVISPPrimartAddy, $projectalpha) or die(mysql_error());
$row_rsVISPPrimartAddy = mysql_fetch_assoc($rsVISPPrimartAddy);
$totalRows_rsVISPPrimartAddy = mysql_num_rows($rsVISPPrimartAddy);

//echo "13<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsACCI = sprintf("SELECT invoicetraxr.InvoiceSerial, invoicetraxr.acci_RecID, invoicetraxr.TotalDue, invoicetraxr.AmountPaid, invoicetraxr.PaymentDue, invoicetraxr.Finalised, invoicetraxr.PaidWhen, invoicetraxr.AmountCredited, accountinfo.AccountName, accountinfo.SysopID, accountinfo.sfStartTime, accountinfo.sfCycle_Upload, accountinfo.sfCycle_Download, accountinfo.sfCycle_Mins, accountinfo.BillingDate, accountinfo.gPassword FROM invoicetraxr inner join accountinfo on invoicetraxr.acci_RecID = accountinfo.RecID where invoicetraxr.RecID = %d",$nTraxrID);
$rsACCI = mysql_query($query_rsACCI, $projectalpha) or die(mysql_error());
$row_rsACCI = mysql_fetch_assoc($rsACCI);
$totalRows_rsACCI = mysql_num_rows($rsACCI);

//echo "14<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsClientAddy = sprintf("SELECT acci_addresses.ContactName, acci_addresses.Street1, acci_addresses.Street2, acci_addresses.Country, acci_addresses.`State`, acci_addresses.Postcode, acci_addresses.Suburb FROM acci_addresses inner join invoicetraxr on acci_addresses.AccI_RecID = invoicetraxr.acci_RecID where acci_addresses.Checked <> 0 and invoicetraxr.RecID = %d  limit 1,1",$nTraxrID);
$rsClientAddy = mysql_query($query_rsClientAddy, $projectalpha) or die(mysql_error());
$row_rsClientAddy = mysql_fetch_assoc($rsClientAddy);
$totalRows_rsClientAddy = mysql_num_rows($rsClientAddy);

//echo "15<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsCurInv = sprintf("SELECT invoiceout.RecID, invoiceout.AccI_RecID, invoiceout.AmountDue as Due, invoiceout.GSTCharged as GST, invoiceout.AmountPaid as Paid, invoiceout.PaidWhen as pWhen, invoiceout.TraxrID, invoiceout.AmountRefunded as Cred, invoiceout.GSTRefunded as cGST, plantypes.CatNo, plantypes.Description  FROM invoiceout inner join plantypes on invoiceout.ptRecID = plantypes.RecID where invoiceout.TraxrID = %d",$nTraxrID);
$rsCurInv = mysql_query($query_rsCurInv, $projectalpha) or die(mysql_error());
$totalRows_rsCurInv = mysql_num_rows($rsCurInv);

//echo "16<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsTTLInvoice = sprintf("SELECT sum(invoiceout.AmountDue) as sDue,sum( invoiceout.GSTCharged ) as sGST, sum(invoiceout.AmountDue) - sum( invoiceout.GSTCharged ) as TotalDebit, sum(invoiceout.AmountPaid) as  sPaid, sum(invoiceout.AmountRefunded) + sum(invoiceout.GSTRefunded) as Credit,sum(invoiceout.AmountRefunded) as sCred, sum(invoiceout.GSTRefunded) as scGST,(sum(invoiceout.AmountDue) + sum( invoiceout.GSTCharged )) - sum(invoiceout.AmountPaid) - (sum(invoiceout.AmountRefunded) + sum(invoiceout.GSTRefunded)) as Total FROM invoiceout inner join plantypes on invoiceout.ptRecID = plantypes.RecID WHERE invoiceout.TraxrID = %d",$nTraxrID);
$rsTTLInvoice = mysql_query($query_rsTTLInvoice, $projectalpha) or die(mysql_error());
$totalRows_rsTTLInvoice = mysql_num_rows($rsTTLInvoice);

//echo "17<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsPaid = sprintf("SELECT invoicetraxr.InvoiceSerial, invoicetraxr.TotalDue, invoicetraxr.AmountPaid, invoicetraxr.PaymentDue, invoicetraxr.AmountCredited FROM invoicetraxr WHERE invoicetraxr.PaymentDue < '%s' and (invoicetraxr.TotalDue + invoicetraxr.GSTDue) > (invoicetraxr.AmountPaid + invoicetraxr.AmountCredited) and invoicetraxr.acci_RecID = '%d'",$row_rsACCI['PaymentDue'],$row_rsACCI['acci_RecID']);
$rsPaid = mysql_query($query_rsPaid, $projectalpha) or die(mysql_error());
$totalRows_rsPaid = mysql_num_rows($rsPaid);

//echo "18<BR>";
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsTTLPre = sprintf("SELECT sum(invoicetraxr.TotalDue) as SMDue, sum(invoicetraxr.AmountPaid) as SMPaid, Sum(invoicetraxr.AmountCredited) as SMCRED, sum(invoicetraxr.TotalDue) - sum(invoicetraxr.AmountPaid) - Sum(invoicetraxr.AmountCredited) + Sum(invoicetraxr.GSTDue) as OutStanding FROM invoicetraxr WHERE invoicetraxr.PaymentDue <= '%s' and (invoicetraxr.TotalDue + invoicetraxr.GSTDue) > (invoicetraxr.AmountPaid + invoicetraxr.AmountCredited) and invoicetraxr.acci_RecID = '%d'",$row_rsACCI['PaymentDue'],$row_rsACCI['acci_RecID']);
$rsTTLPre = mysql_query($query_rsTTLPre, $projectalpha) or die(mysql_error());
$row_rsTTLPre = mysql_fetch_assoc($rsTTLPre);
$totalRows_rsTTLPre = mysql_num_rows($rsTTLPre);

$Grandtotal = ++$row_rsTTLInvoice['sDue'];
$Grandtotal =  ++$row_rsTTLInvoice['sGST'];
$Grandtotal =  --$row_rsCurInv['Paid'] ;
$Grandtotal =  --$row_rsCurInv['Cred'] ;
$Grandtotal =  --$row_rsCurInv['cGST']  ;
$Grandtotal = ++$row_rsTTLPre['OutStanding'];
//$Grandtotal = --$row_rsTTLPre['SMPaid'];
//$Grandtotal = ++$row_rsTTLPre['SMDue'];
//$Grandtotal = ++$row_rsTTLPre['SMDue'];
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Tax Invoice - <?php echo $row_rsVISPa['Description']; ?> rbn: <?php echo $RBN; ?> - Invoice Number <?php echo $_GET['nTraxrID']; ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	color: #000000;
	font-family: Tahoma, Trebuchet MS, Verdana, Lucida Console;
}
body {
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style18 {color: #000000}
.style28 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"}
.style29 {font-weight: bold; color: #000000; font-size: 9px;}
.style31 {font-size: 10px; font-weight: bold;}
.style33 {font-weight: bold; font-size: 12px; color: #FF3333;}
.style34 {font-weight: bold; font-size: 12px; color: #0066FF;}
.style37 {font-size: 12px}
.style38 {color: #0000FF}
.style39 {font-size: 9px}
.style40 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; color: #000000; }
.style41 {font-size: 24px; font-weight: bold;}
.style42 {
	color: #FFFFFF;
	font-weight: bold;
}
.style43 {
	color: #FFFFFF;
	font-size: 18px;
}
.style44 {color: #FFFFFF}
.style50 {font-size: 24; font-weight: bold; color: #CC0000; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; }
.style51 {font-size: 14px; color: #FFFFFF; }
.style54 {font-size: 10px}
.style55 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-size: 10px; }
.style56 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-weight: bold; font-size: 10px; }
a:link {
	color: #FFFFFF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #CCCC33;
}
a:hover {
	text-decoration: none;
	color: #FF9966;
}
a:active {
	text-decoration: none;
	color: #FF6600;
}
.style57 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-size: 9px; }
.style60 {color: #FF0000; font-weight: bold; }
.style61 {color: #0066FF}
.style62 {color: #3399FF}
.style63 {color: #FF0000}
.style65 {color: #FFFFFF; font-size: 11px; }
.style66 {font-size: 11px}
.style67 {color: #FF0000; font-size: 11px; }
.style68 {color: #3399FF; font-size: 11px; }
.style69 {font-size: 18px}
.style70 {font-weight: bold; color: #339933;}
.style73 {font-size: 16px}
.style75 {font-size: 36px}
.style76 {
	font-size: 16;
	font-weight: bold;
}
.style78 {color: #FFFFFF; font-size: 26px; }
-->
</style>
</head>

<body>
<table width="98%"  border="0" align="center">
  <tr>
    <td bgcolor="#0000CC" class="style28">&nbsp;</td>
  </tr>
</table>
<table width="81%"  border="0" align="center">
  <tr>
    <td bordercolor="#CCCCCC" bgcolor="#0066FF" class="style28">&nbsp;</td>
  </tr>
</table>
<table width="96%"  border="0" align="center" >
  <tr>
    <td width="12%" class="style28"><div align="center" class="style29">
      <p>Brought to you by          <img src="http://www.projectalpha.com.au/images/idents/ep_ident.jpg" width="100" height="100"><br>
        Exitstencil Press P/L Networks<br>
        A.C.N 096 867 775
      </p>
      </div></td>
    <td colspan="2" class="style28"><div align="center">
      <h1 class="style18">Tax Invoice<br>
        <?php echo $row_rsVISPa['Description']; ?><br>
          <span class="style37"><?php if (empty($row_rsVISPa['ACN'])) {
		  									if (empty($row_rsVISPa['ABN'])) {
												echo "No RBN on file";
												} else {
												echo 'ABN: ';
												echo $row_rsVISPa['ABN'];
												}
											} else {
												echo 'ACN: ';
												echo $row_rsVISPa['ACN'];
												} ?>		  </span>      </h1>
      <h3 class="style18">          Invoice Number <span class="style38"><?php echo $_GET['nTraxrID']; ?></span> </h3>
    </div></td>
    <td width="29%" rowspan="4" class="style28"><div align="center"><?php if (empty($row_rsVISPa['LogoURL'])) { 
																				echo 'no logo specified';
																			} else {
																				echo '<img src=\"';
																				echo $row_rsVISPa['LogoURL']; 
																				echo '\" width=\"182\" height=\"211\">';
																				}	?></div>          </td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td colspan="2" class="style28"><div align="center"><span class="style18 style39"><?php echo $row_rsVISPPrimartAddy['Street1']; ?>, <?php echo $row_rsVISPPrimartAddy['Street2']; ?>, <?php echo $row_rsVISPPrimartAddy['Suburb']; ?>, <?php echo $row_rsVISPPrimartAddy['State']; ?>, <?php echo $row_rsVISPPrimartAddy['Postcode']; ?>, <?php echo $row_rsVISPPrimartAddy['Country']; ?></span></div></td>
  </tr>
  <tr>
    <td height="47" class="style28"><h4 align="right" class="style18"><strong>Account name:</strong></h4></td>
    <td colspan="2" class="style28"><blockquote>
      <h4 class="style18"><strong><?php echo $row_rsACCI['AccountName']; ?></strong></h4>
    </blockquote></td>
  </tr>
  <tr>
    <td class="style28"><div align="right"><span class="style18"><strong>Address:</strong></span></div></td>
    <td colspan="2" class="style40"><blockquote>
      <p align="center" class="style56"><?php echo $row_rsClientAddy['ContactName']; ?><br>
        <?php echo $row_rsClientAddy['Street1']; ?><br>
        <?php echo $row_rsClientAddy['Street2']; ?><br>
        <?php echo $row_rsClientAddy['Suburb']; ?>, <?php echo $row_rsClientAddy['State']; ?>,<br> 
        <?php echo $row_rsClientAddy['Postcode']; ?>, <?php echo $row_rsClientAddy['Country']; ?></p>
    </blockquote></td>
  </tr>
  <tr>
    <td class="style28"><div align="right"><span class="style18"><strong>Next Statement: </strong></span></div></td>
    <td width="41%" class="style40"><blockquote>
      <p><?php echo $row_rsACCI['BillingDate']; ?></p>
    </blockquote></td>
    <td width="18%" class="style40"><blockquote>
      <p>Server Timestamp: </p>
    </blockquote></td>
    <td class="style40"><div align="right">
      <blockquote>
        <p><?php echo $row_rsSysNow['SysNow']; ?></p>
      </blockquote>
    </div></td>
  </tr>
  <tr>
    <td class="style28"><h3 align="right" class="style18"><strong>Invoice Due:</strong></h3></td>
    <td class="style28"><blockquote>
      <h3 class="style18"><strong><?php echo $row_rsACCI['PaymentDue']; ?></strong></h3>
    </blockquote></td>
    <td class="style28"><blockquote>
      <p><span class="style40"><strong>Invoice Total:</strong></span></p>
    </blockquote></td>
    <td class="style28"><div align="right">
      <blockquote>
        <p><span class="style40"><strong>$ <?php echo $Grandtotal; ?></strong></span></p>
      </blockquote>
    </div></td>
  </tr>
  <tr>
    <td class="style40">&nbsp;</td>
    <td colspan="3" class="style40">&nbsp;</td>
  </tr>
  <tr>
    <td class="style40"><div align="right"></div></td>
    <td colspan="3" class="style40">&nbsp;</td>
  </tr>
  <tr>
    <td class="style28"><span class="style18"><strong>Email Sent To: </strong></span></td>
    <td colspan="3" class="style40">&nbsp;</td>
  </tr>
  <tr>
    <td class="style40">&nbsp;</td>
    <td colspan="3" class="style40">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="4" class="style28" bgcolor="#000099"><div align="center" class="style18"><span class="style41 style44">Outstanding Invoice Summary <br>
    (6 Months History)</span></div></td>
  </tr>
  <tr>
    <td colspan="4" class="style28"><div align="center" class="style18">
      <table width="87%" border="0">
        <tr bgcolor="#000000">
          <td width="25%"><div align="center" class="style44"><strong>Invoice Number</strong></div></td>
          <td width="12%"><div align="center" class="style44"><strong>Total Due</strong></div></td>
          <td width="15%"><div align="center" class="style44"><strong>Amount Paid </strong></div></td>
          <td width="22%"><div align="center" class="style44"><strong>Payment Due </strong></div></td>
          <td width="12%"><div align="center" class="style44"><strong>Credited</strong></div></td>
          <td width="14%"><div align="center" class="style44"><strong>Status</strong></div></td>
        </tr>
	<?php
		while ($row_rsPaid = mysql_fetch_assoc($rsPaid)) {
		?>
        <tr bgcolor="#FFCC99" class="style40">
          <td><span class="style60"><?php echo $row_rsPaid['InvoiceSerial']; ?></span></td>
          <td><div align="right"><span class="style60">$ <?php echo sprintf('%01.2f',$row_rsPaid['TotalDue']); ?></span></div></td>
          <td><div align="right"><span class="style60">$ <?php echo sprintf('%01.2f',$row_rsPaid['AmountPaid']); ?></span></div></td>
          <td><div align="center"><span class="style60"><?php echo $row_rsPaid['PaymentDue']; ?></span></div></td>
          <td><div align="right"><span class="style60">$ <?php echo sprintf('%01.2f',$row_rsPaid['AmountCredited']); ?></span></div></td>
			
          <td><div align="center" class="style60"><?php if ($row_rsPaid['AmountPaid'] == 0) {
		  													echo "Outstanding";
														} else {
															if ($row_rsPaid['AmountPaid'] <  $row_rsPaid['TotalDue']) {
																echo "Partially Paid";
										
															} else {
																echo "Finalised";
															} } ?>  </div></td>
        </tr>
		<?php } ?>
        <tr bgcolor="#000000">
          <td><div align="right" class="style44">
            <blockquote>
              <p><strong>Totals: </strong></p>
            </blockquote>
          </div></td>
          <td><div align="right" class="style44">
            <div align="right"><strong>$ <?php echo sprintf('%01.2f',$row_rsTTLPre['SMDue']); ?></strong></div>
          </div></td>
          <td><div align="right" class="style44">
            <div align="right"><strong><strong>$ </strong><?php echo sprintf('%01.2f',$row_rsTTLPre['SMPaid']); ?></strong></div>
          </div></td>
          <td><div align="right"><span class="style44"></span></div></td>
          <td><div align="right" class="style44">
            <div align="right"><strong>$ <?php echo sprintf('%01.2f',$row_rsTTLPre['SMCRED']); ?></strong></div>

          </div></td>
          <td><div align="right"><span class="style44"></span></div></td>
        </tr>
      </table>
    </div></td>
  </tr>
  <tr>
    <td class="style40">&nbsp;</td>
    <td colspan="3" class="style40">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="4" class="style28" bgcolor="#0000CC"><h1 align="center" class="style44">Additonal Invoice Itinerary</h1></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="4" class="style28"><div align="center" class="style18">
      <table width="95%" border="0">
        <tr bgcolor="#000000" class="style56">
          <td width="99"><span class="style65">Product Code </span></td>
          <td width="257"><span class="style65">Description</span></td>
          <td width="72"><div align="right" class="style63 style66">Due</div></td>
          <td width="83"><div align="right" class="style67">GST Due</div></td>
          <td width="94"><div align="right" class="style66"><span class="style44">Paid</span></div></td>
          <td width="121"><div align="right" class="style66"><span class="style44">Paid When </span></div></td>
          <td width="76"><div align="right" class="style68">Credited</div></td>
          <td width="86"><div align="right" class="style68">GST Credited </div></td>
        </tr>
        <tr bgcolor="#CCCC99" class="style56">
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
        </tr>
		<?php while ($row_rsCurInv = mysql_fetch_assoc($rsCurInv)) { ?>
        <tr bgcolor="#CCCC99" class="style56">
          <td><span class="style66"><?php echo $row_rsCurInv['CatNo']; ?></span></td>
          <td><span class="style66"><?php echo $row_rsCurInv['Description']; ?></span></td>
          <td><div align="right" class="style67">$ <?php echo sprintf('%01.2f',$row_rsCurInv['Due']); ?></div></td>
          <td><div align="right" class="style67">$ <?php echo sprintf('%01.2f',$row_rsCurInv['GST']); ?></div></td>
          <td><div align="right" class="style66">$ <?php echo sprintf('%01.2f',$row_rsCurInv['Paid']); ?></div></td>
          <td><div align="right" class="style66"><?php echo $row_rsCurInv['pWhen']; ?></div></td>
          <td><div align="right" class="style68">$ <?php echo sprintf('%01.2f',$row_rsCurInv['Cred']); ?></div></td>
          <td><div align="right" class="style68">$ <?php echo sprintf('%01.2f',$row_rsCurInv['cGST']); ?></div></td>
        </tr>
		<?php } ?>
        <tr bgcolor="#CCCC99" class="style56">
          <td><span class="style66"></span></td>
          <td><span class="style66"></span></td>
          <td><div align="right"><span class="style63"><span class="style66"></span></span></div></td>
          <td><div align="right"><span class="style63"><span class="style66"></span></span></div></td>
          <td><div align="right"><span class="style66"></span></div></td>
          <td><div align="right"><span class="style66"></span></div></td>
          <td><div align="right"><span class="style61"><span class="style62"><span class="style66"></span></span></span></div></td>
          <td><div align="right"><span class="style62"><span class="style66"></span></span></div></td>
        </tr>
		<?php while ($row_rsTTLInvoice = mysql_fetch_assoc($rsTTLInvoice)) { ?>
        <tr bgcolor="#000000" class="style56">
          <td>&nbsp;</td>
          <td><div align="right" class="style66 style44">
            <blockquote><span class="style44">Totals:</span></blockquote>
          </div></td>
          <td><div align="right" class="style66 style44"><span class="style63"><span class="style18"><span class="style44">$ <?php echo sprintf('%01.2f',$row_rsTTLInvoice['sDue']); ?></span></span></span></div></td>
          <td><div align="right" class="style66 style44"><span class="style63"><span class="style18"><span class="style44">$ <?php echo sprintf('%01.2f',$row_rsTTLInvoice['sGST']); ?></span></span></span></div></td>
          <td><div align="right" class="style66 style44"><span class="style18"><span class="style44">$ <?php echo sprintf('%01.2f',$row_rsTTLInvoice['sPaid']); ?></span></span></div></td>
          <td><div align="right"><span class="style18"><span class="style44"><span class="style66"><span class="style44"></span></span></span></span></div></td>
          <td><div align="right" class="style66 style44"><span class="style61"><span class="style62"><span class="style18"><span class="style44">$ <?php echo sprintf('%01.2f',$row_rsTTLInvoice['sCred']); ?></span></span></span></span></div></td>
          <td><div align="right" class="style66 style44"><span class="style62"><span class="style18"><span class="style44">$<?php echo sprintf('%01.2f',$row_rsTTLInvoice['scGST']); ?></span></span></span></div></td>
        </tr>
		<?php } ?>
      </table>
    </div></td>
  </tr>
  <tr>
    <td class="style40">&nbsp;</td>
    <td class="style40">&nbsp;</td>
    <td class="style40">&nbsp;</td>
    <td class="style40">&nbsp;</td>
  </tr>
  <tr>
    <td class="style28"><div align="right"><span class="style33">Credited:</span></div></td>
    <td class="style28"><blockquote>
      <p><span class="style33">$ <?php echo $row_rsTTLInvoice['Credit']; ?></span></p>
    </blockquote></td>
    <td class="style28"><div align="right" class="style69"><span class="style70">Previously Outstanding: </span></div></td>
    <td class="style28"><div align="right" class="style69"><span class="style70">$ <?php echo sprintf('%01.2f',$row_rsTTLPre['OutStanding']); ?></span></div></td>
  </tr>
  <tr>
    <td class="style28"><div align="right"><span class="style34">Debited:</span></div></td>
    <td class="style28"><blockquote>
      <p><span class="style34">$ <?php echo $row_rsTTLInvoice['TotalDebit']; ?></span></p>
    </blockquote></td>
    <td rowspan="2" class="style41"><div align="right"><span class="style50"> Total Now Due:</span></div></td>
    <td rowspan="2" class="style41"><div align="right"><span class="style50">$ <?php echo $Grandtotal; ?></span></div></td>
  </tr>
  <tr>
    <td class="style37 style28"><div align="right"><strong>Paid:</strong></div></td>
    <td class="style37 style28"><blockquote>
      <p><strong>$ <?php echo $row_rsTTLInvoice['sPaid']; ?></strong></p>
    </blockquote></td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr bgcolor="#0033CC">
    <td colspan="2" bgcolor="#0033CC" class="style28"><h3 class="style42">Payment Options </h3></td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28"><div align="center"><img src="http://www.projectalpha.com.au/images/icons/top_left.gif" width="247" height="69"></div></td>
    <td colspan="2" class="style28"><p>Account name: <strong>Exitstencil Press Pty. Ltd.<br>
        <br>
        </strong>        Account Number: <strong>121481196</strong> <br>
    BSB: <strong>633000 <br>
    <br>
    </strong>Transaction Reference: <strong>PA<?php echo $_GET['nTraxrID']; ?>INV<br>
    </strong> </p>    </td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr bgcolor="#0000CC">
    <td colspan="2" class="style28"><strong><span class="style43">Support and Accounts </span></strong></td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="4" class="style28"><blockquote>
      <p align="justify">        <span class="style37">Many thanks for choosing <?php echo $row_rsVISPa['Description']; ?> for your service provider. <?php echo $row_rsVISPa['Description']; ?> is sponsored onto Exitstencil Press Network by a supporting companies. This relationship ensures that you the end subscriber or purchaser of products on project alphas sales channels are ensure only the best quality and highest standard of telecommunication and products. Should you wish to contact Exitstencil Press Pty Ltd. Regarding the billing of this Invoice please be advised that you can contact us a the <?php echo $row_rsVISPa['Description']; ?> Head quarters:</span><br>
        <br>
        <strong>Accounts:</strong>
        Contact  through email at 
        <a href="mailto:accounts@ep.net.au">accounts@ep.net.au</a> or +61-02-9797-9365 (EST GMT 10+).<br>
        <strong>Support:</strong> Contact support through email at <a href="mailto:support@ep.net.au">support@ep.net.au</a> or +61-02-9517-9140 (EST GMT 10+).<br>
        <br>
        <span class="style37">        Below is the contact details of your visp, if you should have any question surrounding further purchases or changes to services you should contact them first. <?php echo $row_rsVISPa['Description']; ?> is the person who maintains your account. For any Password Changes, Ammendments to services, Reseller Intiative, General Queries or most general internet queries you will have to take up with your visp, we urge you to contact Exitstencil Press Pty ltd only when the ViSP&copy; is unable to fix or answer your question.</span></p>
    </blockquote></td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
  </tr>
  <tr bgcolor="#0033FF">
    <td colspan="3" class="style28"><strong><span class="style78"><?php echo $row_rsVISPa['Description']; ?> Contact Details </span></strong></td>
    <td class="style28"><div align="center"><strong><span class="style51">sponsored by <?php echo $row_rsSponsor['Description']; ?></span></strong></div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style28 style54">&nbsp;</td>
    <td class="style55">&nbsp;</td>
    <td class="style55">&nbsp;</td>
    <td class="style55">&nbsp;</td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style28"><div align="right" class="style76">Contact Admin/Sysop: </div></td>
    <td class="style28"><blockquote class="style76"><a href="mailto:<?php echo $row_rsSysop['Email']; ?>"><?php echo $row_rsSysop['Firstname']; ?> <?php echo $row_rsSysop['Surname']; ?></a><br>
      Sec. Level: <?php echo $row_rsSysop['SecurityLevel']; ?>%<br>
    </blockquote></td>
    <td class="style28"><div align="right" class="style76">Website:</div></td>
    <td class="style28"><blockquote class="style76">
      <?php if (!empty($row_rsVISPa['Realm'])) { ?>
      <a href="http://www.<?php echo $row_rsVISPa['Realm']; ?>/">http://www.<?php echo $row_rsVISPa['Realm']; ?>/</a>      <?php } ?> 
    </blockquote></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style28" ><div align="right" class="style41">Sales:</div></td>
    <td class="style28"><div align="left" class="style31">
      <blockquote>
        <p align="left"><span class="style73">Phone: <?php echo $row_rsVISPPhnA['ContactName']; ?><br>
            <?php echo $row_rsVISPPhnA['PhoneNumber']; ?> <?php if (!empty($row_rsVISPPhnA['Extension'])) { ?> ext. <?php echo $row_rsVISPPhnA['Extension'];  } ?></span><br>
            <br>
            Address: 
            <?php echo $row_rsVISPAddyA['ContactName']; ?><br>
            <?php echo $row_rsVISPAddyA['Street1']; ?><br>
            <?php echo $row_rsVISPAddyA['Street2']; ?><br>
            <?php echo $row_rsVISPAddyA['Suburb']; ?>, <?php echo $row_rsVISPAddyA['State']; ?>, <?php echo $row_rsVISPAddyA['Postcode']; ?>.
            <br>
            <br>
          </p>
      </blockquote>
    </div></td>
    <td class="style28"><div align="right" class="style41 style75">ABN, RBN, ACN:</div></td>
    <td class="style28"><div align="left" class="style31 style75">
      <blockquote>
        <p><?php echo $RBN; ?></p>
      </blockquote>
    </div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style28"><div align="right" class="style41">Accounts:</div></td>
    <td class="style28"><blockquote class="style54">
      <p align="left"><strong><span class="style73">Phone: <?php echo $row_rsVISPPhnB['ContactName']; ?><br>
          <?php echo $row_rsVISPPhnB['PhoneNumber']; ?> <?php if (!empty($row_rsVISPPhnB['Extension'])) { ?> ext. <?php echo $row_rsVISPPhnB['Extension'];  } ?></span><br>
          <br>
          Address: 
          <?php echo $row_rsVISPAddyb['ContactName']; ?><br>
          <?php echo $row_rsVISPAddyb['Street1']; ?><br>
          <?php echo $row_rsVISPAddyb['Street2']; ?><br>
          <?php echo $row_rsVISPAddyb['Suburb']; ?>, <?php echo $row_rsVISPAddyb['State']; ?>, <?php echo $row_rsVISPAddyb['Postcode']; ?>.
          <br>
          <br>
          </strong></p>
    </blockquote></td>
    <td class="style28"><div align="right" class="style41"><strong>Support:</strong></div></td>
    <td class="style28"><div align="right" class="style31">
      <blockquote>
        <div align="left">
          <p align="left"><span class="style73">Phone: <?php echo $row_rsVISPPhnC['ContactName']; ?><br>
              <?php echo $row_rsVISPPhnC['PhoneNumber']; ?> <?php if (!empty($row_rsVISPPhnC['Extension'])) { ?> ext. <?php echo $row_rsVISPPhnC['Extension'];  } ?></span><br>
              <br>
              Address: 
              <?php echo $row_rsVISPAddyC['ContactName']; ?><br>
              <?php echo $row_rsVISPAddyC['Street1']; ?><br>
              <?php echo $row_rsVISPAddyC['Street2']; ?><br>
              <?php echo $row_rsVISPAddyC['Suburb']; ?>, <?php echo $row_rsVISPAddyC['State']; ?>, <?php echo $row_rsVISPAddyC['Postcode']; ?>.
              <br>
              <br>
          </p>
        </div>
      </blockquote>
    </div></td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style56">&nbsp;</td>
    <td class="style56">&nbsp;</td>
    <td class="style56">&nbsp;</td>
    <td class="style56">&nbsp;</td>
  </tr>
  <tr bgcolor="#0033CC">
    <td colspan="4" class="style28"><strong><span class="style43">Terms and Conditions </span></strong></td>
  </tr>
  <tr>
    <td colspan="4" class="style57"><div align="justify">
      <blockquote>&nbsp;
        </blockquote>
    </div>      <blockquote>
      <p align="justify">1.0 EXITSTENCIL PRESS PTY LTD SERVICES/NETWORK<br>
          1.1 In addition to this Agreement, you accept <?php echo $row_rsVISPa['Description']; ?> &amp; Exitstencil Press Pty Ltd may update &amp; or introduce;<br>
    (a) Acceptable Usage Policy. <br>
    (b) Internet Fair Use Policy. <br>
    (c) Any other policy required in response to laws &amp;/or changes in Internet regulation.<br>
    1.2 (Service Availability) Availability or continuity of services at all times or in all areas is not guaranteed. <?php echo $row_rsVISPa['Description']; ?> will endeavor to provide a reliable, trouble free service but due to factors outside of our control, our service delivery may be impacted. Should a third-party fail to continue supporting your geographical area, we may terminate or suspend the services at any time.<br>
    1.3 (Connection) <?php echo $row_rsVISPa['Description']; ?> may: <br>
          (a) Apply an idle time, being a predetermined period, in accordance with the selected plan of inactivity across your connection. When an idle time is reached the connection is released; <br>
          (b) Apply a session time, being a predetermined period, in accordance with the selected plan of time or data permitted in one connection. When a session time is reached the connection is released; <br>
          (c) Suspend your connection to a point of presence or the Internet without giving you notice in the event of network failure or maintenance, to investigate any complaint of illegal behavior or abuse, or if required by a law enforcement agency; <br>
          (d) &amp; update the services, point of presence numbers or any other features at any time without notice.<br>
          1.4 (Variation) <?php echo $row_rsVISPa['Description']; ?> reserve the right to amend this Agreement &amp; any other agreements applicable to the services contracted at any time. Where variations have a negative impact on the customer, <?php echo $row_rsVISPa['Description']; ?> will provide a minimum of 14 days advance notice. <?php echo $row_rsVISPa['Description']; ?> will notify through electronic media including email &amp; website notices. It is your responsibility to keep abreast of these changes that may impact on you.<br>
          1.5 (Security) Dolphin Communication&rsquo;s policy is to maintain privacy &amp; confidentiality in all our communications with our customers unless otherwise advised. The exception to this is when required by law to disclose such as, but not limited to, a judicial order. You acknowledge &amp; agree that <?php echo $row_rsVISPa['Description']; ?> have no responsibility &amp; assure no liability for such acts or occurrences.<br>
          1.6 (Expiry of Fixed Term Agreements) At the completion of a fixed term agreement or plan for any Service, <?php echo $row_rsVISPa['Description']; ?> will continue providing the service on a month-to-month basis until termination.<br>
          1.7 (Bundling) If you subscribe to multiple services or are receiving a bonus service such as web hosting, we will bundle all services into a single account. A default on any Service may lead to termination of this Agreement &amp; thereby all of your services may be discontinued.<br>
          2.0 INTERNET SERVICES.<br>
          2.1 (Access) You are responsible for the use &amp; conduct of the account supplied through this Agreement. Any user of this account must abide by all Dolphin Communication&rsquo;s Internet active policies at time of use.<br>
          2.2 (Online Services) Through the Internet Service you will have access to products, services &amp; information. <?php echo $row_rsVISPa['Description']; ?> cannot warrant any of this information &amp; recommend a buyer beware approach. In accessing these sites you take full responsibility for any charges incurred &amp; indemnify <?php echo $row_rsVISPa['Description']; ?> from any claim for such product, service or information. <br>
          2.3 (Support) In supporting the Internet Service, you will have access to support primarily through an automated ticketing system accessed via email to support@ep.net.au .Should your support request be for configuring access to our service then our helpdesk will take your call during published support hours. <?php echo $row_rsVISPa['Description']; ?> can only support the Internet configuration aspect of the service. Any hardware, networking or operating system issues are not supported by <?php echo $row_rsVISPa['Description']; ?> &amp; should be taken up with the appropriate vendors.<br>
          2.4 (Usage) Usage must comply with our Acceptable Use Policy, current version available through the www.ep.net.au website. <br>
2.5 (IP Address) <?php echo $row_rsVISPa['Description']; ?> will provide you with a dynamic or static IP address, at our discretion unless otherwise stipulated in the sales contract, each time you use the Service. You have right to use IP addresses assigned during the course of Service but no right to ownership.<br>
3.0 EQUIPMENT <br>
3.1 (Installation) Any equipment provided by <?php echo $row_rsVISPa['Description']; ?> is supplied on a self-install basis unless otherwise stipulated in the sales contract. You will bear any costs incurred by third-party services that may be required in connection with the installation of the equipment to your premises. An example of such a cost would be additional wiring of a telephone point on your premises.<br>
4.0 YOUR OBLIGATIONS<br>
4.1 (Billing) You are responsible for ensuring that <?php echo $row_rsVISPa['Description']; ?> is kept informed of any change to your legal name, address &amp; telephone number.<br>
4.2 (Fees &amp; Charges) You must pay Existencil Press all fees, charges &amp; taxes for which you are liable. Our fees &amp; charges may be varied at any time by providing 14 days notice to you. tax invoices, receipts &amp; statements will be delivered by email. An administration fee may be applicable for mailing documentation.<br>
4.3 (Payment) You must pay all fees, charges &amp; taxes (including any goods &amp; services tax) by the due date. Accounts outstanding may have services crippled or suspended. During such time you will still be liable for payment &amp; charges incurred. <?php echo $row_rsVISPa['Description']; ?> may also at its&rsquo; discretion charge a late fee, vary your services or terminate this Agreement for failure to pay by the due date. Reconnection upon payment of any outstanding amounts may incur a reconnection fee. You will be responsible for any debt collection costs we incur.<br>
4.4 (GST) <?php echo $row_rsVISPa['Description']; ?> will charge the GST for all our services at the prescribed rate at the time of charging.<br>
4.5 (Software) Any software provided by <?php echo $row_rsVISPa['Description']; ?> must be used in accordance with license terms &amp; conditions attached to the software. <?php echo $row_rsVISPa['Description']; ?> does not warrant or support this software unless a warranty or support is stipulated in the sale contract. <br>
4.6 (Intellectual Property) Logos, trademarks of <?php echo $row_rsVISPa['Description']; ?> &amp; its partners remain the property of their respective owners. You may not publish or use, without our prior written consent, any trademark, trade name, logo or service mark of <?php echo $row_rsVISPa['Description']; ?>, or its partners.<br>
4.7 (Capacity) You guarantee that you are 18 years old or older.<br>
4.8 (Unauthorized Use) You must maintain confidentiality of user names, passwords &amp; account information. You must immediately notify <?php echo $row_rsVISPa['Description']; ?> in the event of any unauthorized use of the Service. Multiple concurrent log-ons are not permitted. You must not use the services for unauthorized access of any systems or networks connected to the Internet.<br>
4.9 (Agreement) You will be liable for any subsequent services ordered from this master account. Such orders require your master password &amp; will be added to this account.<br>
5.0 Exitstencil Press Pty Ltd Banks through Bendigo bank and westpac and will automatically debit your credit card &amp;/or bank account for the duration of the service &amp;/or supply of product(s). It is your responsibility to make sure sufficient credit or funds are available at time of drawing.<br>
  5.1 Where payment can not be processed through a bank account or credit card, it is your responsibility to arrange an alternative payment method within 7 days from our invoice.<br>
  5.2 To change credit card details, enquire or resolve disputes please forward any correspondence to accounts@ep.net.au . Or by writing to us at 4 Trinity Ave, Millers Point NSW 2000.</p>
    </blockquote>    </td>
  </tr>
  <tr>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28">&nbsp;</td>
    <td class="style28"><div align="right"><img src="http://www.projectalpha.com.au/images/icons/genpa.gif" width="136" height="36"></div></td>
  </tr>
</table>
<table width="88%"  border="0" align="center" bordercolor="#FFFFFF">
  <tr bgcolor="#000000">
    <td bordercolor="#FBFBF9"  class="style28"><div align="right" class="style44">
      <p align="center">This tax invoice was generated by project alpha &copy; 2005 - http://www.projectalpha.com.au/ <br>
        <span class="style54"><strong><a href="http://www.projectalpha.com.au/cwan/invoices/viewinvoice.php?nTraxrID=<?php echo $_GET['nTraxrID']; ?>&gPass=<?php echo $gPass; ?>&nVirtualID=<?php echo $nVirtualID; ?>">This invoice is located at the following URL http://www.projectalpha.com.au/cwan/invoices/viewinvoice.php?nTraxrID=<?php echo $_GET['nTraxrID']; ?>&amp;gPass=<?php echo $gPass; ?>&amp;nVirtualID=<?php echo $nVirtualID; ?></a></strong></span></p>
      </div></td>
  </tr>
</table>
<p class="style28">&nbsp;</p>
</body>
</html>
<?php
mysql_free_result($rsVISPa);

mysql_free_result($rsSysNow);

mysql_free_result($rsSponsor);

mysql_free_result($rsSysop);

mysql_free_result($rsVISPAddyA);

mysql_free_result($rsVISPAddyb);

mysql_free_result($rsVISPAddyC);

mysql_free_result($rsVISPPhnA);

mysql_free_result($rsVISPPhnB);

mysql_free_result($rsVISPPhnC);

mysql_free_result($rsVISPPrimartAddy);

mysql_free_result($rsACCI);

mysql_free_result($rsClientAddy);

mysql_free_result($rsCurInv);

mysql_free_result($rsTTLInvoice);

mysql_free_result($rsPaid);

mysql_free_result($rsTTLPre);
?>
