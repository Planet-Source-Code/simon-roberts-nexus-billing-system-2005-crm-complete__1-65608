<?php require_once('../connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}
mysql_select_db($database_projectalpha, $projectalpha);
$query_STATS = "SELECT distinct Count(distinct virtualisp.RecID) as CntVisp, STDDEV(virtualisp.JoiningFee) as STDJN, AVG(virtualisp.JoiningFee) as AVGJN FROM virtualisp WHERE virtualisp.CreatedBy_SysopID = 1 ";
$STATS = mysql_query($query_STATS, $projectalpha) or die(mysql_error());
$row_STATS = mysql_fetch_assoc($STATS);
$totalRows_STATS = mysql_num_rows($STATS);

mysql_select_db($database_projectalpha, $projectalpha);
$query_sysops = sprintf("SELECT Count(distinct sysops.RecID) as CntSYS FROM virtualisp inner join sysops on sysops.VirtualID = virtualisp.RecID WHERE virtualisp.CreatedBy_SysopID = %d ",$_SESSION['SysopID']);
$sysops = mysql_query($query_sysops, $projectalpha) or die(mysql_error());
$row_sysops = mysql_fetch_assoc($sysops);
$totalRows_sysops = mysql_num_rows($sysops);

mysql_select_db($database_projectalpha, $projectalpha);
$query_stats3 = sprintf("SELECT Count(distinct plantypes.RecID) as PLNCNT FROM virtualisp inner join plantypes on plantypes.VirtualID = virtualisp.RecID WHERE virtualisp.CreatedBy_SysopID = %d ",$_SESSION['SysopID']);
$stats3 = mysql_query($query_stats3, $projectalpha) or die(mysql_error());
$row_stats3 = mysql_fetch_assoc($stats3);
$totalRows_stats3 = mysql_num_rows($stats3);

mysql_select_db($database_projectalpha, $projectalpha);
$query_yourvisps = sprintf("SELECT distinct virtualisp.RecID , Count(distinct plantypes.RecID) as ProdCount, Count(distinct invoicetraxr.RecID) as InvCount,  Count(distinct sysops.RecID) as SysopsCount, virtualisp.Description FROM virtualisp inner join plantypes ON plantypes.VirtualID = virtualisp.RecID inner join invoicetraxr on invoicetraxr.VirtualID = plantypes.VirtualID inner join sysops on invoicetraxr.VirtualID = sysops.VirtualID WHERE virtualisp.CreatedBy_SysopID = %d GROUP BY virtualisp.RecID",$_SESSION['SysopID']);
$yourvisps = mysql_query($query_yourvisps, $projectalpha) or die(mysql_error());


?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>View Your Visps</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	color: #FF9966;
}
body {
	background-color: #CCCCCC;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-image: url(../images/backgrnd/woodtile2.PNG);
}
a {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: medium;
	color: #993333;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.style1 {
	font-size: x-large;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style3 {font-size: small}
.style5 {font-size: small; color: #993333; }
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; font-size: x-small; }
.style9 {color: #FFFFFF}
-->
</style></head>

<body>

<?php
include("top.php3");
?>
<table width="73%"  border="0" align="center">
  <tr bgcolor="#CCCCCC">
    <td width="14%"><div align="center"><img src="../images/idents/ep_ident.jpg" width="100" height="100"></div></td>
    <td width="38%"><span class="style1">Your Virtual ISPs utilising the Global ViSP&copy; Network.</span></td>
    <td width="48%"><span class="style1"><span class="style5">Total ViSP Signed: <?php echo $row_STATS['CntVisp']; ?><br>
STD Network Joining Fee: $ <?php echo sprintf("%01.2f",$row_STATS['STDJN']); ?><br>
AVG Network Joining Fee: $ <?php echo sprintf("%01.2f",$row_STATS['AVGJN']); ?><br>
Total Sysops Within Network: <?php echo $row_sysops['CntSYS']; ?><br>
Total Products Within Network: <?php echo $row_stats3['PLNCNT']; ?></span></span></td>
  </tr>
  <tr bgcolor="">
    <td rowspan="2" bgcolor="#CCCCCC"></td>
    <td height="109" colspan="2" bgcolor="#FFFFFF"><table width="99%"  border="0" align="center">
      <tr bgcolor="#000000" class="style8">
        <td width="49%"><span class="style8">Company/Registered Business Name </span></td>
        <td width="16%"><div align="right"><span class="style8">Total Products </span></div></td>
        <td width="17%"><div align="right">Total Invoices </div></td>
        <td width="18%"><div align="right"><span class="style8">Total Sysops </span></div></td>
      </tr>
	  <?php
	  while ($row_yourvisps = mysql_fetch_assoc($yourvisps)) { ?>
      <tr bgcolor="#336633" class="style8">
        <td><span class="style9"><?php echo $row_yourvisps['Description']; ?></span></td>
        <td><div align="right" class="style9"><?php echo $row_yourvisps['ProdCount']; ?></div></td>
        <td><div align="right" class="style9"><?php echo $row_yourvisps['InvCount']; ?></div></td>
        <td><div align="right" class="style9"><?php echo $row_yourvisps['SysopsCount']; ?></div></td>
      </tr>
	  <?php } ?>
      <tr  bgcolor="#666666" class="style8">
        <td>&nbsp;</td>
        <td><div align="right"></div></td>
        <td><div align="right"></div></td>
        <td><div align="right"></div></td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
<?php
mysql_free_result($STATS);

mysql_free_result($sysops);

mysql_free_result($stats3);

mysql_free_result($yourvisps);
?>
