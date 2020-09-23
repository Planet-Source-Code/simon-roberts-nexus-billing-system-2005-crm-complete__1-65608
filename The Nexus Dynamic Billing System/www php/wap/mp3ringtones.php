<?php
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
<?php include("http://www.projectalpha.com.au/wap/wapcss.php?wapartist=0&pagefrom=default"); ?>
</head>

<body>
<div align="center">
  <p><img src="../images/icons/punkringtons.gif" width="200" height="100"></p>
  <p><span class="style1">These Ringtone have been reduce way below cd quality this will not really affect you on a mobile phone due to the speaker size. WARNING do not play on convential audio systems, may cause hearing damage.</span></p>
  <p class="style2 style2">The Crass</p>
  <p><span class="style3"><a href="../audio/ringtones/TheCrass/Crass%20-%20Chairman%20Of%20The%20Bored.mp3">Crass - Chairman Of The Bored.mp3</a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Do%20They%20Owe%20Us%20A%20Living.mp3">Crass - Do They Owe Us A Living.mp3</a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Fight%20War,%20Not%20Wars.mp3">Crass - Fight War, Not Wars.mp3</a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Fuck%20All%20Government.mp3">Crass - Fuck All Government.mp3</a> <br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20General%20Bacardi.mp3">Crass - General Bacardi.mp3</a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Mickey%20Mouse%20Is%20Dead.mp3">Crass - Mickey Mouse Is Dead.mp3</a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Mother%20Earth.mp3">Crass - Mother Earth.mp3 </a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Punk%20Is%20Dead.mp3">Crass - Punk Is Dead.mp3 </a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Reality%20Asylum.mp3">Crass - Reality Asylum.mp3 </a><br>
      <a href="../audio/ringtones/TheCrass/Crass%20-%20Reality%20Whitewash.mp3">Crass - Reality Whitewash.mp3 </a></span><a href="../audio/ringtones/TheCrass/Crass%20-%20Reality%20Whitewash.mp3">- </a>
  <em>Warning Blasphemy to catholism</em><br>
  <a href="../audio/ringtones/TheCrass/Crass%20-%20Shaved%20Women.mp3">Crass - Shaved Women.mp3 </a><br>
  <a href="../audio/ringtones/TheCrass/Crass%20-%20Smash%20The%20Mac.mp3">Crass - Smash The Mac.mp3 </a><br>
  <a href="../audio/ringtones/TheCrass/Crass%20-%20Upright%20Citizen.mp3">Crass - Upright Citizen.mp3 </a><br>
  <a href="../audio/ringtones/TheCrass/Crass%20-%20White%20Punks%20On%20Hope.mp3">Crass - White Punks On Hope.mp3 </a><br>
  <a href="../audio/ringtones/TheCrass/Crass%20-%20You%27ve%20Got%20Big%20Hands.mp3">Crass - You've Got Big Hands.mp3 </a> </p>
</div>
</body>
</html>

<?php } ?>