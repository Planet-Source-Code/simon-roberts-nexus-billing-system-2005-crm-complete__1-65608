<?php require_once('Connections/projectalpha.php'); ?>
<?php
	error_reporting(0);

	include(sprintf('http://www.sydney-escort-lounge.com.au/botredir.php?incdude=%s',$incdude));	
	
	mysql_select_db($database_projectalpha, $projectalpha);
	$query_rsret = "SELECT * FROM `_pa_botredirect`";
	$rsret = mysql_query($query_rsret, $projectalpha) or die(mysql_error());
	$totalRows_rsret = mysql_num_rows($rsret);
	
	while ($row_rsret = mysql_fetch_assoc($rsret))
	{
		if (strstr(" " . $incdude, $row_rsret['bottype'])) {
		//print $row_rsret['bottype'] . " = " . $incdude . " | " .  '<br>';		
		}
//		if (strstr(" " . $incdude, $row_rsret['bottype']) != 0) 
		if (strstr(" " . $incdude, $row_rsret['bottype'])) 
		{
			
			if ($row_rsret['actiontype'] == 'link') 
			{
				echo sprintf("<a href=\"%s\">[%s] %s </a>",$row_rsret['link'],$row_rsret['bottagid'],$row_rsret['linktext']);
			} else {
				echo $row_rsret['xHTML'];
			}
			
		}
	}

mysql_free_result($rsret);
?>
