<?php require_once('../Connections/projectalpha.php'); ?>
<?php
mysql_select_db($database_projectalpha, $projectalpha);
$query_rsInvItn = sprintf("SELECT invoiceout.AmountDue, invoiceout.GSTCharged, invoiceout.PaymentDue, invoiceout.AmountPaid, invoiceout.PaidWhen, invoiceout.TotalDue, invoiceout.Description FROM invoiceout WHERE invoiceout.TraxrID = %d",$nTraxrID);
$rsInvItn = mysql_query($query_rsInvItn, $projectalpha) or die(mysql_error());
$totalRows_rsInvItn = mysql_num_rows($rsInvItn);

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsInvPayments = sprintf("SELECT invout_payment.Amount, invout_payment.GST, invout_payment.Sub, invout_payment.WhenPaid, invout_payment.TotalPaid, invoiceout.Description  FROM invoiceout, invout_payment WHERE  invoiceout.TraxrID = %d AND invout_payment.InvOut_RecID = invoiceout.RecID",$nTraxrID);
$rsInvPayments = mysql_query($query_rsInvPayments, $projectalpha) or die(mysql_error());
$totalRows_rsInvPayments = mysql_num_rows($rsInvPayments);

mysql_select_db($database_projectalpha, $projectalpha);
$query_ttlInv = sprintf("SELECT sum(invoiceout.AmountDue) as ttlDue, sum(invoiceout.GSTCharged) as ttlGST, sum(invoiceout.AmountPaid) as ttlPaid, sum(invoiceout.TotalDue) as ttltotaldue FROM invoiceout WHERE invoiceout.TraxrID = %d",$nTraxrID);
$ttlInv = mysql_query($query_ttlInv, $projectalpha) or die(mysql_error());
$row_ttlInv = mysql_fetch_assoc($ttlInv);
$totalRows_ttlInv = mysql_num_rows($ttlInv);

mysql_select_db($database_projectalpha, $projectalpha);
$query_rsAcci = sprintf("SELECT accountinfo.AccountName FROM accountinfo, invoicetraxr WHERE invoicetraxr.acci_RecID =  accountinfo.RecID and invoicetraxr.InvoiceSerial = %d",$nTraxrID);
$rsAcci = mysql_query($query_rsAcci, $projectalpha) or die(mysql_error());
$row_rsAcci = mysql_fetch_assoc($rsAcci);
$totalRows_rsAcci = mysql_num_rows($rsAcci);

if(!session_id()){
  session_start();
}
if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view client base</a>";
	exit;
}


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
.style3 {font-size: 10px}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FF3333;
	font-weight: bold;
}
.style13 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
}
body,td,th {
	color: #0000CC;
}
.style14 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"}
.style16 {
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
	font-weight: bold;
	color: #FFFFFF;
}
.style17 {color: #FFFFFF}
.style22 {font-size: 12px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; }
.style23 {font-size: 12px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; color: #FFFFFF; }
.style24 {font-weight: bold}
.style25 {font-size: 12px; font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; color: #FFFFFF; font-weight: bold; }
.style28 {
	font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console";
	font-weight: bold;
	color: #FF9933;
	font-size: 14px;
}
.style30 {color: #0000CC}
.style35 {font-family: Tahoma, "Trebuchet MS", Verdana, "Lucida Console"; font-size: 14px; font-weight: bold; }
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<table width="770"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td width="18%"><h3 class="style28"><span class="style30">Invoice Number:</span></h3></td>
    <td width="82%"><span class="style35"><?php echo $_GET['nTraxrID']; ?></span></td>
  </tr>
  <tr>
    <td><span class="style35">Clients Name:</span></td>
    <td><span class="style35"><?php echo $row_rsAcci['AccountName']; ?></span></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><span class="style1">Items on the Invoice </span></td>
  </tr>
  <tr>
    <td colspan="2"><table width="97%" border="0" align="center">
      <tr bgcolor="#000000">
        <td width="172"><div align="center"><strong><span class="style23">Subscripton or Product</span></strong></div></td>
        <td width="121"><div align="center"><strong><span class="style23">Total Due </span></strong></div></td>
        <td width="131"><div align="center"><strong><span class="style23">Paid so Far </span></strong></div></td>
        <td width="133"><div align="center"><strong><span class="style23">GST/Tax Charged </span></strong></div></td>
        <td width="133"><div align="center"><strong><span class="style23">Payment Due </span></strong></div></td>
      </tr>

	   <?php 
      while ($row_rsInvItn = mysql_fetch_assoc($rsInvItn)) {
	
  ?>
      <tr bgcolor="#CCCCCC">
        <td><span class="style22"><?php echo $row_rsInvItn['Description']; ?></span></td>
        <td><div align="right"><span class="style22">$ <?php echo sprintf('%01.2f',$row_rsInvItn['TotalDue']); ?></span></div></td>
        <td><div align="right"><span class="style22">$ <?php echo sprintf('%01.2f',$row_rsInvItn['AmountPaid']); ?></span></div></td>
        <td><div align="right"><span class="style22">$ <?php echo sprintf('%01.2f',$row_rsInvItn['GSTCharged']); ?></span></div></td>
        <td><div align="center"><span class="style22"><?php echo $row_rsInvItn['PaymentDue']; ?></span></div></td>
      </tr>
	  <?php } ?>
      <tr bgcolor="#000000">
        <td><div align="right" class="style17"><strong><span class="style22">Totals:</span></strong></div></td>
        <td><div align="right" class="style23 style24">
          <div align="right">$ <?php echo sprintf('%01.2f',$row_ttlInv['ttltotaldue']); ?></div>
        </div></td>
        <td><div align="right" class="style25">
          <div align="right">$ <?php echo sprintf('%01.2f',$row_ttlInv['ttlPaid']); ?></div>
        </div></td>
        <td><div align="right" class="style25">
          <div align="right">$ <?php echo sprintf('%01.2f',$row_ttlInv['ttlGST']); ?></div>
        </div></td>
        <td><div align="right"></div></td>
      </tr>
    </table>
    <div align="center"></div></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><span class="style1">Payment made so far on this invoice </span></td>
  </tr>
  <tr>
    <td colspan="2"><table width="99%"  border="0" cellspacing="1" cellpadding="0" align="center">
      <tr bgcolor="#000000">
        <td width="45%" ><span class="style16">Description</span></td>
        <td width="26%"><div align="right" class="style14 style17"><strong>Total Payment </strong></div></td>
        <td width="29%"><div align="center" class="style16">Time Stamp of Payment </div></td>
        </tr>
	   <?php 
      while ($row_rsInvPayments = mysql_fetch_assoc($rsInvPayments)) {
	
  ?>
      <tr class="style8" bgcolor="#CCCCCC">
        <td><span class="style14"><?php echo $row_rsInvPayments['Description']; ?></span></td>
        <td><div align="right" class="style14">$ <?php echo sprintf('%01.2f',$row_rsInvPayments['TotalPaid']); ?></div></td>
        <td><div align="center" class="style14"><?php echo $row_rsInvPayments['WhenPaid']; ?></div></td>
        </tr>
	  <?php
}
 ?>
      <tr class="style8" bgcolor="#000000">
        <td><span class="style14"></span></td>
        <td><span class="style14"></span></td>
        <td><span class="style14"></span></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><div align="center"><a href="login.php"><img src="../images/icons/cd-rom.jpg" width="32" height="32" /></a><br />      
    <a href="login.php"><span class="style13">Back To Main</span> </a></div></td>
    <td><div align="center"><a href="login.php"><br />
        </a></div></td>
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
mysql_free_result($rsInvItn);

mysql_free_result($rsInvPayments);

mysql_free_result($ttlInv);

mysql_free_result($rsAcci);


?>
