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
$query_rsInv = sprintf("SELECT accountinfo.AccountName, invoicetraxr.RecID, invoicetraxr.InvoiceSerial, invoicetraxr.TotalDue+invoicetraxr.GSTDue as AmtDue, invoicetraxr.AmountPaid, invoicetraxr.PaymentDue, invoicetraxr.PaidWhen, invoicetraxr.AmountCredited FROM accountinfo, invoicetraxr WHERE invoicetraxr.acci_RecID = accountinfo.RecID and accountinfo.SysopID = %d and ((invoicetraxr.TotalDue+invoicetraxr.GSTDue) -  invoicetraxr.AmountPaid - invoicetraxr.AmountCredited) > 1 ORDER BY invoicetraxr.InvoiceSerial",$_SESSION['SysopID']);
$rsInv = mysql_query($query_rsInv, $projectalpha) or die(mysql_error());

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
.style7 {color: #FF3333}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #FF3333;
	font-weight: bold;
}
.style9 {color: #0066CC}
.style10 {color: #0099FF}
.style12 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #FF6633; }
.style13 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
}
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<table width="770"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td><h3 class="style1">Active Invoices On Your Client Base </h3></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2"><table width="99%"  border="0" cellspacing="1" cellpadding="0" align="center">
      <tr>
        <td width="17%"><div align="center" class="style3 style2 style7">
          <div align="left"><strong>Account Name </strong></div>
        </div></td>
        <td width="14%"><div align="center" class="style8">
          <p align="center">Amount Due<br />
            (INC Tax)
          </p>
          </div></td>
        <td width="16%"><div align="center" class="style8">
          <div align="center">Amount Paid<br />
              (INC Tax)
          </div>
        </div></td>
        <td width="17%"><div align="center" class="style8">
          <div align="center">Payment Due</div>
        </div></td>
        <td width="17%"><div align="center" class="style8">
          <div align="center">Payment Made </div>
        </div></td>
        <td width="19%"><div align="center" class="style8">Credited<br />
          (INC Tax) </div></td>
      </tr>
	   <?php 
      while ($row_rsInv = mysql_fetch_assoc($rsInv)) {
	
  ?>
      <tr class="style8">
        <td><div align="left" class="style10"><span class="style2"><span class="style3"><?php echo $row_rsInv['InvoiceSerial']; ?> - <?php echo $row_rsInv['AccountName']; ?></span></span></div></td>
        <td><div align="right" class="style10"><span class="style2"><?php echo $row_rsInv['AmtDue']; ?></span></div></td>
        <td><div align="center" class="style10">
              <div align="right"><span class="style12"><a href="<?php echo sprintf('inv_payments.php?nTraxrID=%d',$row_rsInv['RecID']); ?>"><?php echo $row_rsInv['AmountPaid']; ?></a></span></div>
          </div></td>
        <td><div align="center" class="style10">
          <div align="center"><span class="style2"><?php echo $row_rsInv['PaymentDue']; ?></span></div>
        </div></td>
        <td><div align="right" class="style10">
          <div align="center"><?php echo $row_rsInv['PaidWhen']; ?></div>
        </div></td>
        <td><div align="center" class="style10">
          <div align="right"><span class="style2"><?php echo $row_rsInv['AmountCredited']; ?></span></div>
        </div></td>
      </tr>
	  <?php
;
	} ?>
      <tr class="style8">
        <td><div align="center"><span class="style2"><span class="style3"><span class="style9"></span></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><span class="style9"></span></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><span class="style9"></span></span></span></div></td>
        <td><div align="center"><span class="style2"><span class="style3"><span class="style9"></span></span></span></div></td>
        <td><span class="style9"></span></td>
        <td><div align="center"><span class="style2"><span class="style3"><span class="style9"></span></span></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="center"><a href="login.php"><img src="../images/icons/cd-rom.jpg" width="32" height="32" /><br />
        <span class="style13">Back To Main</span> </a></div></td>
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
mysql_free_result($rsInv);
?>
