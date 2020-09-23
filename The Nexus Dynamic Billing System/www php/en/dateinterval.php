<?php require_once('../connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}
if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view client base</a>";
	exit;
}
print $Save;

if (!empty($usersubmit)) {
	if (($txtStartDt <> $_SESSION['StartForcast']) or ($txtEndDt <> $_SESSION['EndForcast'])) {
		mysql_select_db($database_projectalpha, $projectalpha);
		$sqla = sprintf("update sysops set StartForcast = '%s', EndForcast = '%s' where RecID = '%d'",$txtStartDt, $txtEndDt, $_SESSION['SysopID']);
		$rsSQLa = mysql_query($sqla, $projectalpha) or die(mysql_error());
		$sqlb = sprintf("update sysops set StartForcast =  DateCreated where RecID = '%d' and StartForcast < DateCreated", $_SESSION['SysopID']);
		$rsSQLb = mysql_query($sqlb, $projectalpha) or die(mysql_error());
		$sqlc = sprintf("Select StartForcast, EndForcast from sysops where RecID = '%d'", $_SESSION['SysopID']);
		$rsSQLc = mysql_query($sqlc, $projectalpha) or die(mysql_error());
		$row_rsSQLc = mysql_fetch_assoc($rsSQLc);
		$_SESSION['StartForcast'] = $row_rsSQLc['StartForcast'];
		$_SESSION['EndForcast'] = $row_rsSQLc['EndForcast'];
	}
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
.style14 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-size: 16px;
	color: #0000FF;
	font-weight: bold;
}
-->
</style></head>

<body>
<?php
include("top.php3");
?>
<form name="DatesQuery" id="DatesQuery" method="post" action="<?php print "$PHP_SELF"; ?>">
<table width="65%"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td colspan="2"><span class="style14">Set forcast date interval - set the dates your reporting go between, remembering to specify at least a valid date if not inclusive of time. ie. yyyy-mm-dd Hh:Nn:Ss</span></td>
  </tr>
  <tr>
    <td width="27%">&nbsp;</td>
    <td width="73%">&nbsp;</td>
  </tr>
  <tr class="style14">
    <td><div align="right">Start Forcast Datetime:</div></td>
    <td>
      <input name="txtStartDt" type="text" id="txtStartDt" value="<?php echo sprintf('%s',$_SESSION['StartForcast']) ?>" size="50" />
    </td>
  </tr>
  <tr class="style14">
    <td><div align="right">End Forcast Datetime:</div></td>
    <td>
      <input name="txtEndDt" type="text" id="txtEndDt" value="<?php echo sprintf('%s',$_SESSION['EndForcast']) ?>" size="50" />
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
      <div align="left">
        <p>&nbsp;        </p>	
              <input type="Submit" name="usersubmit" value="Submit &amp; Save Dates" />
        <p>&nbsp;        </p>
    </div></td>
  </tr>
</table>
</form>
<?php
include("bottom.php3");
?>
</body>
</html>
