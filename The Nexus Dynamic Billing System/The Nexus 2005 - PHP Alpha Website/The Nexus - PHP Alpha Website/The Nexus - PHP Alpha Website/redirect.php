<?php require_once('Connections/Epwebdev.php'); ?>
<?php
mysql_select_db($database_Epwebdev, $Epwebdev);
$click = "update rdFiles set NumberOfHits = (NumberOfHits + 1) where ID = $nFileID";
$excclick = mysql_query($click, $Epwebdev) or die(mysql_error());

mysql_select_db($database_Epwebdev, $Epwebdev);
$query_redirect = "SELECT rdFiles.Title FROM rdFiles where ID = $nFileID";
$redirect = mysql_query($query_redirect, $Epwebdev) or die(mysql_error());
$row_redirect = mysql_fetch_assoc($redirect);
$totalRows_redirect = mysql_num_rows($redirect);
?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><?php 
  	echo sprintf("Redirect: %s",$row_redirect['Title']); 
  ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>

<frameset rows="28,476"  cols="*" framespacing="3" frameborder="yes" border="1" bordercolor="#9966FF">
  <frame noresize scrolling="no" src="<?php 
  	echo sprintf("redirect_header.php?nFileID=%s",$nFileID); 
  ?>" />
  <frame src="<?php 
  	echo sprintf("redirect_footer.php?nFileID=%s",$nFileID); 
  ?>" />
</frameset>
<noframes><body>
</body></noframes>
</html>
