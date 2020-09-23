<?php require_once('Connections/Epwebdev.php'); ?>
<?php
mysql_select_db($database_Epwebdev, $Epwebdev);
$query_redirect = "SELECT `rdFiles`.`HeaderHTML`, `rdFiles`.`HeaderBG-Colour`, `rdFiles`.`HeaderFG-Colour` FROM rdFiles where ID = $nFileID";
$redirect = mysql_query($query_redirect, $Epwebdev) or die(mysql_error());
$row_redirect = mysql_fetch_assoc($redirect);
$totalRows_redirect = mysql_num_rows($redirect);
?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Trebuchet MS, Tahoma, Arial;
	font-size: 10px;
	color: #<?php echo $row_redirect['HeaderFG-Colour']; ?>;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #<?php echo $row_redirect['HeaderBG-Colour']; ?>;
}
-->
</style></head>

<body>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#<?php echo $row_redirect['HeaderBG-Colour']; ?>">
    <td width="15%" align="center" valign="top" bgcolor="#<?php echo $row_redirect['HeaderBG-Colour']; ?>"><img src="images/icons/genpa.gif" width="136" height="36"></td>
    <td width="85%" align="center" valign="middle" bgcolor="#<?php echo $row_redirect['HeaderBG-Colour']; ?>"><div align="center"><?php echo $row_redirect['HeaderHTML']; ?></div></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
