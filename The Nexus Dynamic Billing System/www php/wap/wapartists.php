
<?php

	require('../Connections/Epwebdev.php');
	
 	function iif($clause, $truestate, $falsestate) 
	{
		if ($clause == true) {
			return $truestate;
		} else {
			return $falsestate;
		}
	}
	
	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_wapprofile = sprintf("SELECT wapartists_index.wawmail_header, wapartists_index.wawmail_splitbar, wapartists_index.waEmail, wapartists_index.waProfileName, wapartists_index.waURLComment, wapartists_index.waURL, wapartists_index.waTLImg, wapartists_index.waTImg, wapartists_index.waTRImg, wapartists_index.waBlImg, wapartists_index.waBImg, wapartists_index.waBRImg, wapartists_index.waHeaderComment, wapartists_index.waFooterComment, wapartists_index.waBackground FROM wapartists_index where warid=%s",$wapartist);
	$wapprofile = mysql_query($query_wapprofile, $Epwebdev) or die(mysql_error());
	$row_wapprofile = mysql_fetch_assoc($wapprofile);
	$totalRows_wapprofile = mysql_num_rows($wapprofile);
	
		if ($totalRows_wapprofile == 0) {

		} else {
		
			$profilename = $row_wapprofile['waProfileName'];
			$urlcomment = $row_wapprofile['waURLComment'];
			$aurl = $row_wapprofile['waURL'];
			$hlimg = $row_wapprofile['waTLImg'];
			$hmimg = $row_wapprofile['waTImg'];
			$hrimg = $row_wapprofile['waTRImg'];
			$blimg = $row_wapprofile['waBlImg'];
			$bmimg = $row_wapprofile['waBImg'];
			$brimg = $row_wapprofile['waBRImg'];
			$headertext = $row_wapprofile['waHeaderComment'];
			$footertext = $row_wapprofile['waFooterComment'];
			$background = $row_wapprofile['waBackground'];
			$imgheader = $row_wapprofile['wawmail_header'];
			$splitbar = $row_wapprofile['wawmail_splitbar'];
			$email = $row_wapprofile['waEmail'];				
		
		}


		$xmlcomplant=trim($HTTP_ACCEPT);
	
		$browser=trim($HTTP_USER_AGENT);
	
		$pagetitle = "Wap Artist #$wapartist - $profilename";
	
		$htmlxml ="xmlns=\"http://www.w3.org/1999/xhtml\"";
	
		$doctype = "html PUBLIC \"-//NOKIA//DTD XHTML Mobile +CHTML 1.0//EN\" \"http://www.nokia.com/dtd/xhtml-mp-chtml.dtd\"";
	
	  if (strpos(" " . strtoupper($xmlcomplant),"XML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"XHTML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
		
	  if (strpos(" " . strtoupper($xmlcomplant),"WML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"MP3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"MPEG3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}	
		
	  if (strpos(" " . strtoupper($xmlcomplant),"M4A") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}

	  if (strpos(" " . strtoupper($xmlcomplant),"MP4") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}			
		
	  if (strpos(" " . strtoupper($xmlcomplant),"MPEG4") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}											
	  if (strpos(" " . strtoupper($browser),"R380") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPI") != false)			// Ericsson WapIDE 2.0
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MC21") != false)			// Ericsson MC218
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"AUR ") != false)			// Ericsson R320
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ERIC") != false)			// Ericsson R380
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UP.B") != false)			// UP.Browser
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WINW") != false)			// WinWAP browser
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UPG1") != false)			// UP.SDK 4.0
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UPSI") != false)			// another kind of UP.Browser ??
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"QWAP") != false)			// unknown QWAPPER browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"JIGS") != false)			// unknown JigSaw browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"JAVA") != false)			// unknown Java based browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ALCA") != false)			// unknown Alcatel-BE3 browser (UP based?)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MITS") != false)			// unknown Mitsubishi browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MOT-") != false)			// unknown browser (UP based?)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MY S") != false)           // unknown Ericsson devkit browser ?
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPJ") != false)			// Virtual WAPJAG www.wapjag.de
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"FETC") != false)			// fetchpage.cgi Perl script from www.wapcab.de
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ALAV") != false)			// yet another unknown UP based browser ?
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPA") != false)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MOZI") != false)
			{
			$xmlyes = true;
		} 
	  if (strpos(" " . strtoupper($browser),"NOKI") != false)             // another unknown browser (Web based "Wapalyzer"?)
		{
			$xmlyes = true;
		}
		

if ($xmlyes == true) {
?>
<?php echo sprintf("%s%sxml version=\"1.0\"? encoding=\"utf\-8\"%s%s","<","?","?",">"); ?> 
<!DOCTYPE <?php echo $doctype; ?>> 
<html <?php echo $htmlxml; ?>>
<head>
<title><?php echo $pagetitle; ?></title>
<?php include(sprintf("http://www.projectalpha.com.au/wap/wapcss.php?wapartist=%s&pagefrom=wapartists",$wapartist)); ?>
</head>

<body>
	<?php if ($hmimg != "") { ?>
      <div align="right">
        <p align="center"><img width="100%" height="55" src="<?php echo $hmimg; ?>"/>                    
          <?php } ?>
        <div align="center"><a href="<?php echo $aurl; ?>"><?php echo $urlcomment; ?></a><br />
		<div align="center"><?php echo $headertext; ?></div><br />
        <div align="left">
		  <form name="form1" id="form1" method="post" action="">
 
            <?php include(sprintf("http://www.projectalpha.com.au/wap/waptracks.php?wapartist=%s",$wapartist)); ?>
            <p align="center">              <?php echo $footertext; ?> <br />
	          <a href="wapwmail.php?wapartist=<?php echo $wapartist; ?>">Leave <?php echo $profilename; ?> some cheers!</a>
            </p>
            </p>
          </form>
	    </div>
<p align="center"><strong>There has been <?php 	
	$pagecounted='wapartists';
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
	
mysql_free_result($wapcountr1); ?> Visits </strong></p>
<p align="left">
            <?php if ($blimg != "") { ?>
            <img src="<?php echo $blimg; ?>"/>
            <?php } else { ?>
            <?php } ?>
        </p>
<p align="center">            <?php if ($bmimg != "") { ?>
            <img src="<?php echo $bmimg; ?>"/>
            <?php } else { ?>
            <?php } ?>
        </p>
          <p align="right">            <?php if ($brimg != "") { ?>
            <img src="<?php echo $brimg; ?>"/>
            <?php } else { ?>
            <?php } ?>
          

        
          </p>
</div>
</body>
</html><?php
mysql_free_result($wapprofile);
}
?>
