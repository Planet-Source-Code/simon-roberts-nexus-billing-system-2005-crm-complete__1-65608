<?php require('../Connections/Epwebdev.php'); ?>
<?php


 	function iif($clause, $truestate, $falsestate) 
	{
		if ($clause == true) {
			return $truestate;
		} else {
			return $falsestate;
		}
	}


		if (!empty($sendwmail)) {
		
			if (!empty($waUserMessage) &&
				!empty($waUsername) &&
				!empty($waPhoneNumber)) {
					
												
				if (strpos(" " . strtoupper($waSendAsEmail),"YES") != false) {
							
					require_once('../forum/mainfile.php');
					require_once('../forum/class/xoopsmailer.php');
					
					$adminMessage = sprintf("To %s, 
					
					You have recieved a nominated wmail from %s from the number %s,
					
					|--:[The Message is as follows]:--|		
					
					%s",$profilename,$waUsername,$waPhoneNumber,$waUserMessage);		
					
					$subject = sprintf("wmail from %s",$waUsername);
					
					$XoopsMailer = & getMailer();
					$XoopsMailer->useMail();
					$XoopsMailer->setToEmails($email);
					$XoopsMailer->setFromEmail('simon@projectalpha.com.au');
					$XoopsMailer->setFromName('Project Alpha - Wap Artist wMail');
					$XoopsMailer->setSubject($subject);
					$XoopsMailer->setBody($adminMessage);
					$XoopsMailer->send();
				
				}
				
				
				mysql_select_db($database_Epwebdev, $Epwebdev);
				$query_wapprofile = sprintf("insert into wapartists_usermessages (warid, waUserMessage, waUsername, waPhoneNumber, waDisplayNumber, waSendAsEmail) values ('%s','%s','%s','%s','%s','%s')",$wapartist, $waUserMessage, $waUsername, $waPhoneNumber, iif((strpos(" " . strtoupper($waDisplayNumber),"YES") != false),'Yes','No'), iif((strpos(" " . strtoupper($waSendAsEmail),"YES") != false),'Yes','No'));
				$wapprofile = mysql_query($query_wapprofile, $Epwebdev) or die(mysql_error());
				
				$xmaildone = true;
				$waUserMessage = "";
				$waUsername = "";
						
			} else {
				$xmaildone = false;
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
	
		$pagetitle = "Wap Artist #$wapartist w-Mail $profilename";
	
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
<?php include(sprintf('http://www.projectalpha.com.au/wap/wapcss.php?wapartist=%s&pagefrom=wapwmail',$wapartist)); ?>
</head>
<body>
<img src="<?php echo $imgheader; ?>" width="100%" height="60"></br>
<form name="usersubmit" method="post" action=""><label>Phone:</label> 
      <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="waPhoneNumber" type="text" id="waPhoneNumber" size="30">
      </br>
    <br>
  <label>Message:</label> 
  <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="waUserMessage" type="text" id="waUserMessage" size="40">
  </br>
  <br>
  <label>Alias:</label>
      <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="waUsername" type="text" id="waUsername" size="10">
      </br>
    <br>
      <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="waDisplayNumber" type="checkbox" id="waDisplayNumber" value="Yes" checked>
      <label>Hide Phone Number</label></br>
    <br>
      <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="waSendAsEmail" type="checkbox" id="waSendAsEmail" value="Yes">
<label>Send as Email </label></br>
  </p>
  <p align="center">      <input <?php echo iif(($xmaildone==true),'disabled',''); ?> name="sendwmail" type="submit" id="sendwmail" value="Send wMail to <?php echo $profilename; ?>">
    </p>
</form>
<p align="center"><strong>There has been <?php 
	$pagecounted='wapwmail';
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
 ?> Visits </strong></p>
<img src="<?php echo $splitbar; ?>" width="100%" height="51"></br>
 
  <span class="style2">
<?php 
  
  mysql_select_db($database_Epwebdev, $Epwebdev);
$query_wapprofile = sprintf("SELECT wapartists_usermessages.waSendAsEmail, wapartists_usermessages.waDisplayNumber, wapartists_usermessages.waPhoneNumber, wapartists_usermessages.waUsername, wapartists_usermessages.waUserMessage FROM wapartists_usermessages where wapartists_usermessages.warid = %s order by waServerTime desc limit 30",$wapartist);
$wapprofile = mysql_query($query_wapprofile, $Epwebdev) or die(mysql_error());
$totalRows_wapprofile = mysql_num_rows($wapprofile);


  while ($row_wapprofile = mysql_fetch_assoc($wapprofile)) { 
  ?>
Alias:</span> <i><?php echo $row_wapprofile['waUsername']; ?><br>
</i><span class="style2">Message:</span> <b><i><?php echo $row_wapprofile['waUserMessage']; ?></i></b></br>
<br>
  <?php 
  	if ($row_wapprofile['waDisplayNumber'] == 'No') {
	?>
  <span class="style2">Phone:</span><?php echo $row_wapprofile['waPhoneNumber']; ?></ul></br>
  
  <?php 
  	}
	?>
	<br><hr>
<?php
  	}
}
	?>

</body>
</html>