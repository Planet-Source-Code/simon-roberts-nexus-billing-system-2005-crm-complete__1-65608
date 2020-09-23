<?php require_once('Connections/Epwebdev.php'); ?>
<?php
  //error_reporting(0);
  
		mysql_select_db($database_Epwebdev, $Epwebdev);
		$query_lastpost_id = "select count(EngineID) as numengn from _SearchEngineList ";
		$lastpost_id = mysql_query($query_lastpost_id, $Epwebdev) or die(mysql_error());
		$row_lastpost_id = mysql_fetch_assoc($lastpost_id);
		
		mysql_select_db($database_Epwebdev, $Epwebdev);
		$query_lastposthist = sprintf("select EngineID, EngineName, EngineNormalurl, EngineSearchURL, EngineSpace4What from _SearchEngineList limit %s,1",rand(0,$row_lastpost_id['numengn']-1));
		$lastposthist = mysql_query($query_lastposthist, $Epwebdev) or die(mysql_error());
		$row_lastposthist = mysql_fetch_assoc($lastposthist);
		$totalRows_lastposthist = mysql_num_rows($lastposthist);		
		
		$htmlredirect = $row_lastposthist['EngineSearchURL'];
		$nhurl = $row_lastposthist['EngineNormalurl'];
		$space4what = $row_lastposthist['EngineSpace4What'];
		$engineid = $row_lastposthist['EngineID'];
		
		$hostip=$REMOTE_ADDR;
		$hostname = gethostbyaddr($REMOTE_ADDR);
		$xmlcomplant=trim($HTTP_ACCEPT);
		$browser=trim($HTTP_USER_AGENT);
		
		$sqla = sprintf("insert into _Kill_TSN_Survey (RemoteIP, RemoteHostname, DomainSearched, EngineHopID, xml_html, browser) VALUES('%s','%s','%s','%s','%s','%s')", $hostip, $hostname, $lookfor, $engineid, $xmlcomplant, $browser);

		if (strlen($lookfor) > 0) {
			$htmlredirect = str_replace('<QUEST01 />',str_replace(" ", $space4what, $lookfor), $htmlredirect);
		} else {
			$htmlredirect = $nhurl;
		}
			
		mysql_free_result($lastpost_id);
		mysql_free_result($lastposthist);

		mysql_select_db($database_Epwebdev, $Epwebdev);
		$lastposthist = mysql_query($sqla, $Epwebdev) or die(mysql_error());
	
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<title>Random Search Engine - <?php echo $nhurl; ?></title>
<?php 
include(sprintf("http://www.projectalpha.com.au/meta.php?nexpv01='%s'&nexpv02='%s'",'HTML%20Redirector', 'HTML%20Redirector')); 
?>
<META HTTP-EQUIV="refresh" CONTENT="0;URL=<?php echo sprintf("%s",$htmlredirect); ?>">
<body>
<SCRIPT LANGUAGE=javascript>
setTimeout('window.location.href = "<?php
echo sprintf("%s",$htmlredirect); ?>";',0)   </SCRIPT>
    
</body>
</html>
