<?php require_once('../connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}
if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view sales history</a>";
	exit;
}

mysql_select_db($database_projectalpha, $projectalpha);
$query_b_NumInvoices = sprintf("SELECT Count(*) as bNumInvoices FROM invoiceout inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where (AmountPaid +  AmountRefunded + GSTRefunded) < (AmountDue + GSTCharged) and AmountPaid > 0 and accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ", $_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$b_NumInvoices = mysql_query($query_b_NumInvoices, $projectalpha) or die(mysql_error());
$row_b_NumInvoices = mysql_fetch_assoc($b_NumInvoices);
$totalRows_b_NumInvoices = mysql_num_rows($b_NumInvoices);

mysql_select_db($database_projectalpha, $projectalpha);
$query_sysop = sprintf("SELECT sysops.Username, sysops.Firstname, sysops.Surname FROM sysops WHERE sysops.RecID = %s",$_SESSION['SysopID']);
$sysop = mysql_query($query_sysop, $projectalpha) or die(mysql_error());
$row_sysop = mysql_fetch_assoc($sysop);
$totalRows_sysop = mysql_num_rows($sysop);

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsNow = "SELECT distinct NOW() as ServerTime from accountclass";
$rsNow = mysql_query($query_rsNow, $projectalpha) or die(mysql_error());
$row_rsNow = mysql_fetch_assoc($rsNow);
$totalRows_rsNow = mysql_num_rows($rsNow);

mysql_select_db($database_projectalpha, $projectalpha);
$query_TotalCust = sprintf("SELECT count(accountinfo.RecID) as CustCount FROM accountinfo Where SysopID = %s",$_SESSION['SysopID']);
$TotalCust = mysql_query($query_TotalCust, $projectalpha) or die(mysql_error());
$row_TotalCust = mysql_fetch_assoc($TotalCust);
$totalRows_TotalCust = mysql_num_rows($TotalCust);

mysql_select_db($database_projectalpha, $projectalpha);
$query_visps = sprintf("SELECT virtualisp.RecID, virtualisp.ABN, virtualisp.ACN, virtualisp.Description, virtualisp.Realm, virtualisp.JoiningFee FROM virtualisp WHERE virtualisp.CreatedBy_SysopID = %s and virtualisp.CreationDate >=  '%s' and virtualisp.CreationDate <= '%s' ", $_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$visps = mysql_query($query_visps, $projectalpha) or die(mysql_error());
$totalRows_visps = mysql_num_rows($visps);

mysql_select_db($database_projectalpha, $projectalpha);
$query_accinfo = sprintf("SELECT count(acci_services.RecID) as CntServices, sum(((acci_services.PeriodFee + acci_services.JoiningFee) * 0.1 )+(acci_services.PeriodFee + acci_services.JoiningFee) ) as Total, accountinfo.RecID, accountinfo.AccountName, accountinfo.BillingDate, accountclass.Description FROM accountinfo inner join accountclass on accountinfo.Classification = accountclass.RecID inner join acci_services on acci_services.AccI_RecID = accountinfo.RecID where accountinfo.SysopID = %s and accountinfo.CreationDate >=  '%s' and accountinfo.CreationDate <= '%s' GROUP BY accountinfo.RecID", $_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$accinfo = mysql_query($query_accinfo, $projectalpha) or die(mysql_error());
$totalRows_accinfo = mysql_num_rows($accinfo);

mysql_select_db($database_projectalpha, $projectalpha);
$query_v_Joining = sprintf("SELECT Sum(virtualisp.JoiningFee) as SumJoining FROM virtualisp  where virtualisp.CreatedBy_SysopID = %s and virtualisp.CreationDate >=  '%s' and virtualisp.CreationDate <= '%s' ",$_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$v_Joining = mysql_query($query_v_Joining, $projectalpha) or die(mysql_error());
$row_v_Joining = mysql_fetch_assoc($v_Joining);
$totalRows_v_Joining = mysql_num_rows($v_Joining);

mysql_select_db($database_projectalpha, $projectalpha);
$query_tv_Joining = sprintf("SELECT Sum(JoiningFee) FROM virtualisp where virtualisp.CreatedBy_SysopID = %s and virtualisp.CreationDate >=  '%s' and virtualisp.CreationDate <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);;
$tv_Joining = mysql_query($query_tv_Joining, $projectalpha) or die(mysql_error());
$row_tv_Joining = mysql_fetch_assoc($tv_Joining);
$totalRows_tv_Joining = mysql_num_rows($tv_Joining);

mysql_select_db($database_projectalpha, $projectalpha);
$query_c_NumInvoices = sprintf("SELECT Count(*) as bNumInvoices FROM invoiceout inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID WHERE AmountPaid >= (AmountDue + GSTCharged) and accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$c_NumInvoices = mysql_query($query_c_NumInvoices, $projectalpha) or die(mysql_error());
$row_c_NumInvoices = mysql_fetch_assoc($c_NumInvoices);
$totalRows_c_NumInvoices = mysql_num_rows($c_NumInvoices);

mysql_select_db($database_projectalpha, $projectalpha);
$query_c_NumPO = sprintf("SELECT Count(Distinct POID) as aNumPo FROM acci_services, invoiceout inner join accountinfo on accountinfo.RecID = acci_services.acci_RecID where accountinfo.SysopID = %s and DateCreated >=  '%s' and DateCreated <= '%s' ",$_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$c_NumPO = mysql_query($query_c_NumPO, $projectalpha) or die(mysql_error());
$row_c_NumPO = mysql_fetch_assoc($c_NumPO);
$totalRows_c_NumPO = mysql_num_rows($c_NumPO);

mysql_select_db($database_projectalpha, $projectalpha);
$query_a_TotalSales = sprintf("SELECT sum(invoiceout.AmountDue+invoiceout.GSTCharged) as SumSales, sum(invoiceout.AmountPaid) as SumPaid, sum(invoiceout.GSTCharged) as SumGST,sum(invoiceout.AmountRefunded+GSTRefunded) as SumCredited, sum(plantemplates.PeriodFee +(plantemplates.PeriodFee * 0.1)) as SumCost, sum(invoiceout.AmountDue + invoiceout.GSTCharged -(plantemplates.PeriodFee+(plantemplates.PeriodFee * 0.1))) as SumMargin, AVG(invoiceout.AmountPaid/(plantemplates.PeriodFee+(plantemplates.PeriodFee * 0.1)-(invoiceout.AmountRefunded+GSTRefunded))*100) as AvgGPPerc FROM invoiceout inner join plantypes on invoiceout.ptRecID = plantypes.RecID inner join plantemplates on plantypes.TemplateID = plantemplates.RecID inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$a_TotalSales = mysql_query($query_a_TotalSales, $projectalpha) or die(mysql_error());
$row_a_TotalSales = mysql_fetch_assoc($a_TotalSales);
$totalRows_a_TotalSales = mysql_num_rows($a_TotalSales);

mysql_select_db($database_projectalpha, $projectalpha);
$query_b_TotalSales = sprintf("SELECT sum(invoiceout.AmountDue+invoiceout.GSTCharged) as SumSales, sum(invoiceout.AmountPaid) as SumPaid, sum(invoiceout.GSTCharged) as SumGST,sum(invoiceout.AmountRefunded+GSTRefunded) as SumCredited, sum(plantemplates.PeriodFee +(plantemplates.PeriodFee * 0.1)) as SumCost, sum(invoiceout.AmountDue + invoiceout.GSTCharged -(plantemplates.PeriodFee+(plantemplates.PeriodFee * 0.1))) as SumMargin, AVG(invoiceout.AmountPaid/(plantemplates.PeriodFee+(plantemplates.PeriodFee * 0.1)-(invoiceout.AmountRefunded+GSTRefunded))*100) as AvgGPPerc FROM invoiceout inner join plantypes on invoiceout.ptRecID = plantypes.RecID inner join plantemplates on plantypes.TemplateID = plantemplates.RecID  inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);
$b_TotalSales = mysql_query($query_b_TotalSales, $projectalpha) or die(mysql_error());
$row_b_TotalSales = mysql_fetch_assoc($b_TotalSales);
$totalRows_b_TotalSales = mysql_num_rows($b_TotalSales);

mysql_select_db($database_projectalpha, $projectalpha);
$query_ta_NumInvoices = sprintf("SELECT Count(*) as taNumInvoices FROM invoiceout  inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);
$ta_NumInvoices = mysql_query($query_ta_NumInvoices, $projectalpha) or die(mysql_error());
$row_ta_NumInvoices = mysql_fetch_assoc($ta_NumInvoices);
$totalRows_ta_NumInvoices = mysql_num_rows($ta_NumInvoices);

mysql_select_db($database_projectalpha, $projectalpha);
$query_tb_NumInvoices = sprintf("SELECT Count(*) as bNumInvoices FROM invoiceout inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where (AmountPaid +  AmountRefunded + GSTRefunded) < (AmountDue + GSTCharged) and AmountPaid > 0 and accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);
$tb_NumInvoices = mysql_query($query_tb_NumInvoices, $projectalpha) or die(mysql_error());
$row_tb_NumInvoices = mysql_fetch_assoc($tb_NumInvoices);
$totalRows_tb_NumInvoices = mysql_num_rows($tb_NumInvoices);

mysql_select_db($database_projectalpha, $projectalpha);
$query_tc_NumInvoices = sprintf("SELECT Count(*) as taNumInvoices FROM invoiceout inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where AmountPaid >= (AmountDue + GSTCharged) and accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);
$tc_NumInvoices = mysql_query($query_tc_NumInvoices, $projectalpha) or die(mysql_error());
$row_tc_NumInvoices = mysql_fetch_assoc($tc_NumInvoices);
$totalRows_tc_NumInvoices = mysql_num_rows($tc_NumInvoices);

mysql_select_db($database_projectalpha, $projectalpha);
$query_ta_NumPO = sprintf("SELECT Count(Distinct POID) as aNumPo FROM acci_services inner join accountinfo on accountinfo.RecID = acci_services.acci_RecID where accountinfo.SysopID = %s and acci_services.DateCreated >=  '%s' and acci_services.DateCreated <= '%s' ",$_SESSION['SysopID'],$_SESSION['sysCreated'],$_SESSION['EndForcast']);
$ta_NumPO = mysql_query($query_ta_NumPO, $projectalpha) or die(mysql_error());
$row_ta_NumPO = mysql_fetch_assoc($ta_NumPO);
$totalRows_ta_NumPO = mysql_num_rows($ta_NumPO);

mysql_select_db($database_projectalpha, $projectalpha);
$query_a_NumInvoices = sprintf("SELECT Count(*) as aNumInvoices FROM invoiceout inner join accountinfo on accountinfo.RecID = invoiceout.AccI_RecID where accountinfo.SysopID = %s and invoiceout.Created >=  '%s' and invoiceout.Created <= '%s' ",$_SESSION['SysopID'],$_SESSION['StartForcast'],$_SESSION['EndForcast']);
$a_NumInvoices = mysql_query($query_a_NumInvoices, $projectalpha) or die(mysql_error());
$row_a_NumInvoices = mysql_fetch_assoc($a_NumInvoices);
$totalRows_a_NumInvoices = mysql_num_rows($a_NumInvoices);
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Sales Chart <?php echo $_SESSION['StartForcast']?> to <?php echo $_SESSION['EndForcast'] ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #0066CC;
}
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style3 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; }
.style4 {font-size: 12px}
.style5 {
	font-size: 12px;
	font-weight: bold;
}
.style6 {
	color: #0099FF;
	font-weight: bold;
}
body,td,th {
	color: #FFFFFF;
}
.style8 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style10 {font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 16px; }
.style11 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
}
a:link {
	color: #FFFFFF;
	text-decoration: none;
}
a:visited {
	color: #FFCC33;
	text-decoration: none;
}
a:hover {
	color: #FFFF99;
	text-decoration: none;
}
a:active {
	color: #FF3366;
	text-decoration: none;
}
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<table width="770"  border="0" align="center">
  <tr>
    <td colspan="4"><div align="center">
        <h1><span class="style1">Sysop Sales Report </span></h1>
    </div></td>
    <td rowspan="2"><div align="center"><img src="/images/idents/ep_ident.jpg" width="100" height="100"></div></td>
  </tr>
  <tr>
    <td colspan="4"><div align="center" class="style8">
      <p>Report Generated for [ <?php echo $row_sysop['Username']; ?> ] - <?php echo $row_sysop['Firstname']; ?> <?php echo $row_sysop['Surname']; ?></p>
      </div></td>
  </tr>
  <tr>
    <td><div align="right" class="style5"><span class="style1">Server Time: </span></div></td>
    <td><div align="right" class="style10"><?php echo $row_rsNow['ServerTime']; ?></div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="18%"><div align="center"></div></td>
    <td width="32%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
    <td width="22%">&nbsp;</td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Number of Invoices Generated: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_a_NumInvoices['aNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong><span class="style1">Report Start Date: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $_SESSION['StartForcast']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Partly Paid Invoices:</span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_b_NumInvoices['bNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong><span class="style1">Report End Date: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $_SESSION['EndForcast']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Full Paid Invoices: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_c_NumInvoices['cNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Purchase Orders Generated: </strong></div></td>
    <td><div align="right"><strong><?php echo $row_c_NumPO['aNumPo']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong>Total Sales Value: </strong></div></td>
    <td class="style6"><div align="right">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumSales']); ?></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Total Cost:</strong></div></td>
    <td><div align="right" class="style6">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumCost']); ?></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style5">GST Invoiced: </div></td>
    <td class="style6"><div align="right">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumGST']); ?></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Margin:</strong></div></td>
    <td><div align="right" class="style6">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumMargin']); ?></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right">
      <div align="right" class="style4"><strong>Total Amount Paid: </strong></div>
    </div></td>
    <td class="style6"><div align="right">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumPaid']); ?></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style5"><strong><em>Completed Budget:</em></strong></div></td>
    <td><div align="right" class="style6"><?php echo sprintf("%01.4f",$row_a_TotalSales['AvgGPPerc']); ?>%</div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style5">Total Amount Credited: </div></td>
    <td class="style6"><div align="right">$ <?php echo sprintf("%01.2f",$row_a_TotalSales['SumCredited']); ?></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style5">Resellers, Wholsalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP Joining Fee Earnt: </div></td>
    <td><div align="right" class="style6">$ <?php echo sprintf("%01.2f",$row_v_Joining['SumJoining']); ?></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td>&nbsp;</td>
    <td colspan="3"><div align="center">
      <h3>Resellers, Wholsalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP Signed in this Cycle</h3>
    </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td colspan="5"><table width="95%"  border="1" align="center">
      <tr class="style5">
        <td width="28%"><div align="center"><strong>Group/Company Name </strong></div></td>
        <td width="18%"><div align="center"><strong>ABN</strong></div></td>
        <td width="16%"><div align="center"><strong>ACN</strong></div></td>
        <td width="21%"><div align="center">Realm/Domain</div></td>
        <td width="17%"><div align="center"><strong>Joining Fee </strong></div></td>
      </tr>
	    <?php 
      while ($row_visps = mysql_fetch_assoc($visps)) {
  ?>

      <tr class="style5">
        <td><span class="style1"><?php echo $row_visps['Description']; ?></span></td>
        <td><div align="center"><span class="style1"><?php echo $row_visps['ABN']; ?></span></div></td>
        <td><div align="center"><span class="style1"><?php echo $row_visps['ACN']; ?></span></div></td>
        <td><div align="center"><?php echo $row_visps['Realm']; ?></div></td>
        <td><div align="right"><span class="style1">$ <?php echo sprintf("%01.2f",$row_visps['JoiningFee']); ?></span></div></td>
      </tr>
	  <?php
	  }
	  ?>
      <tr class="style5">
        <td><span class="style1"></span></td>
        <td><span class="style1"></span></td>
        <td><span class="style1"></span></td>
        <td>&nbsp;</td>
        <td><span class="style1"></span></td>
      </tr>
    </table></td>
  </tr>
  <tr class="style1">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td>&nbsp;</td>
    <td colspan="3"><div align="center">
      <h3><strong>Subscribers Signed in this Cycle </strong></h3>
    </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td colspan="5"><table width="95%"  border="1" align="center">
      <tr class="style5">
        <td width="35%"><div align="center">Account Description</div></td>
        <td width="19%"><div align="center">Next Billing Date </div></td>
        <td width="22%"><div align="center"><span class="style1">Classification</span></div></td>
        <td width="11%"><div align="center">Number of Services </div></td>
        <td width="13%"><div align="center"><span class="style1">Total Sale value </span></div></td>
      </tr>
	    <?php 
      while ($row_accinfo = mysql_fetch_assoc($accinfo)) {
  ?>
      <tr class="style5">
        <td><span class="style1"><?php echo $row_accinfo['AccountName']; ?></span></td>
        <td><div align="center"><?php echo $row_accinfo['BillingDate']; ?></div></td>
        <td><div align="center"><span class="style1"><?php echo $row_accinfo['Description']; ?></span></div></td>
        <td><div align="center"><span class="style1"><?php echo $row_accinfo['CntServices']; ?></span></div></td>
        <td><div align="right"><span class="style1">$ <?php echo sprintf("%01.2f",$row_accinfo['Total']); ?> </span></div></td>
      </tr>
	  <?php
	  }
	  ?>
      <tr class="style5">
        <td><span class="style1"></span></td>
        <td>&nbsp;</td>
        <td><span class="style1"></span></td>
        <td><span class="style1"></span></td>
        <td><span class="style1"></span></td>
      </tr>
    </table></td>
  </tr>
  <tr class="style1">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5"><div align="center">
        <h1><span class="style1">Total Overall  Sales </span></h1>
    </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Number of Invoices Generated: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_ta_NumInvoices['taNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong><span class="style1">Totals Start Date: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $_SESSION['sysCreated']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Partly Paid Invoices:</span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_tb_NumInvoices['bNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong><span class="style1">Totals End Date: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $_SESSION['EndForcast']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong><span class="style1">Full Paid Invoices: </span></strong></div></td>
    <td><div align="right"><strong><?php echo $row_tc_NumInvoices['taNumInvoices']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Purchase Order Generated: </strong></div></td>
    <td><div align="right"><strong><?php echo $row_ta_NumPO['aNumPo']; ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong>Total Sales Value: </strong></div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumSales']); ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Total Cost:</strong></div></td>
    <td><div align="right"><strong> $ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumCost']); ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style4"><strong>Total GST Raised: </strong></div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumGST']); ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style4"><strong>Margin:</strong></div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumMargin']); ?></strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right">
        <div align="right" class="style4"><strong>Total Amount Paid: </strong></div>
    </div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumPaid']); ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style5"><strong><em>Completed Budget:</em></strong></div></td>
    <td><div align="right"><strong><?php echo sprintf("%01.4f",$row_b_TotalSales['AvgGPPerc']); ?>%</strong></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style5">Total Amount Credited: </div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_b_TotalSales['SumCredited']); ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style5">Resellers, Wholsalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP Joining Fee Earnt: </div></td>
    <td><div align="right"><strong>$ <?php echo sprintf("%01.2f",$row_tv_Joining['Sum(JoiningFee)']); ?>
      </strong></div>
      <div align="right"></div>
      <div align="right"></div>
      <div align="right"></div>
    <div align="right"></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
    <td><div align="right"></div></td>
    <td><div align="right"></div></td>
  </tr>
  <tr class="style1">
    <td><div align="right" class="style5">Total Customers: </div></td>
    <td><div align="right"><strong><?php echo $row_TotalCust['CustCount']; ?></strong></div></td>
    <td>&nbsp;</td>
    <td><div align="right" class="style5"></div></td>
    <td><div align="center"><a href="login.php"><img src="../images/icons/cd-rom.jpg" width="32" height="32" /><br />
        <span class="style11">Back To Main</span> </a></div></td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>
<?php
mysql_free_result($b_NumInvoices);

mysql_free_result($sysop);

mysql_free_result($rsNow);

mysql_free_result($TotalCust);

mysql_free_result($visps);

mysql_free_result($accinfo);

mysql_free_result($v_Joining);

mysql_free_result($tv_Joining);

mysql_free_result($c_NumInvoices);

mysql_free_result($c_NumPO);

mysql_free_result($a_TotalSales);

mysql_free_result($b_TotalSales);

mysql_free_result($ta_NumInvoices);

mysql_free_result($tb_NumInvoices);

mysql_free_result($tc_NumInvoices);

mysql_free_result($ta_NumPO);

mysql_free_result($a_NumInvoices);
?>
