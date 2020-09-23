<?php

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
			return false;
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
			return true;
		}


		$xmlcomplant=trim($HTTP_ACCEPT);
	
		$browser=trim($HTTP_USER_AGENT);
	
		$pagetitle = "Wap Artist #$wapartist - $profilename";
	
		$htmlxml ="xmlns=\"http://www.w3.org/1999/xhtml\"";
	
		$doctype = "html PUBLIC \"-//NOKIA//DTD XHTML Mobile +CHTML 1.0//EN\" \"http://www.nokia.com/dtd/xhtml-mp-chtml.dtd\"";
	
	  if (strpos(" " . strtoupper($xmlcomplant),"xml") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"xhtml") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
		
	  if (strpos(" " . strtoupper($xmlcomplant),"wml") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"mp3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"mpeg3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}	
		
	  if (strpos(" " . strtoupper($xmlcomplant),"m4a") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}

	  if (strpos(" " . strtoupper($xmlcomplant),"mp4") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}			
		
	  if (strpos(" " . strtoupper($xmlcomplant),"mpeg4") != false)			// Ericsson WAP phones and emulators
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
		
	}
?>