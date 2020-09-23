
<?php require_once('../Connections/Epwebdev.php'); ?>
<?php
	$hostip=$REMOTE_ADDR;
	$hostname = gethostbyaddr($REMOTE_ADDR);

	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_wapcountr1 = sprintf("SELECT wapartists_counteraudit.waraudid FROM wapartists_counteraudit WHERE wapartists_counteraudit.waruIP = '%s' AND wapartists_counteraudit.waruHostname = '%s' AND wapartists_counteraudit.waruHTTPACCEPT = '%s' AND wapartists_counteraudit.waruHTTPUSERAGENT = '%s' AND wapartists_counteraudit.warid = '%s' and wapartists_counteraudit.pagecounted = '%s'",$hostip,$hostname,$xmlcomplant,$browser,$wapartist,$pagecounted);
	$wapcountr1 = mysql_query($query_wapcountr1, $Epwebdev) or die(mysql_error());
	$row_wapcountr1 = mysql_fetch_assoc($wapcountr1);
	$totalRows_wapcountr1 = mysql_num_rows($wapcountr1);

	if ($totalRows_wapcountr1 < 1)
	{
	
		mysql_select_db($database_Epwebdev, $Epwebdev);
		$query_wapcountr2 = sprintf("insert into wapartists_counteraudit (waruIP, waruHostname, waruHTTPACCEPT, waruHTTPUSERAGENT, warid, pagecounted) VALUES('%s','%s','%s','%s','%s','%s')",$hostip,$hostname,$xmlcomplant,$browser,$wapartist,$pagecounted);
		$wapcountr2 = mysql_query($query_wapcountr2, $Epwebdev) or die(mysql_error());
	
	
	}

	mysql_free_result($wapcountr1);
	
	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_wapcountr1 = sprintf("SELECT count(wapartists_counteraudit.waraudid) warnv FROM wapartists_counteraudit WHERE wapartists_counteraudit.warid = '%s' and pagecounted = '%s'",$wapartist, $pagecounted);
	$wapcountr1 = mysql_query($query_wapcountr1, $Epwebdev) or die(mysql_error());
	$row_wapcountr1 = mysql_fetch_assoc($wapcountr1);
	$totalRows_wapcountr1 = mysql_num_rows($wapcountr1);

	?><b><?php echo $row_wapcountr1['warnv']; ?></b><?php
	
mysql_free_result($wapcountr1);
?>
