<?php   if(strpos(" " . strtoupper($HTTP_ACCEPT),"vnd.wap.wml") > 0) {        // Check whether the browser/gateway says it accepts WML.
    //////header("Content-type: text/vnd.wap.wml");
?>
<head>
<title>Free Ringtones? Want a Profile?</title>
<?php
  }
  else {
    $browser=substr(trim($HTTP_USER_AGENT),0,4);
    if($browser=="Noki" ||			// Nokia phones and emulators
      $browser=="Eric" ||			// Ericsson WAP phones and emulators
      $browser=="WapI" ||			// Ericsson WapIDE 2.0
      $browser=="MC21" ||			// Ericsson MC218
      $browser=="AUR " ||			// Ericsson R320
      $browser=="R380" ||			// Ericsson R380
      $browser=="UP.B" ||			// UP.Browser
      $browser=="WinW" ||			// WinWAP browser
      $browser=="UPG1" ||			// UP.SDK 4.0
      $browser=="upsi" ||			// another kind of UP.Browser ??
      $browser=="QWAP" ||			// unknown QWAPPER browser
      $browser=="Jigs" ||			// unknown JigSaw browser
      $browser=="Java" ||			// unknown Java based browser
      $browser=="Alca" ||			// unknown Alcatel-BE3 browser (UP based?)
      $browser=="MITS" ||			// unknown Mitsubishi browser
      $browser=="MOT-" ||			// unknown browser (UP based?)
      $browser=="My S" ||           // unknown Ericsson devkit browser ?
      $browser=="WAPJ" ||			// Virtual WAPJAG www.wapjag.de
      $browser=="fetc" ||			// fetchpage.cgi Perl script from www.wapcab.de
      $browser=="ALAV" ||			// yet another unknown UP based browser ?
      $browser=="Wapa")             // another unknown browser (Web based "Wapalyzer"?)
        {
        //////header("Content-type: text/html.xhtml.xml");
		$xmlyes = true;
    }
    else {
      //////header("Content-type: text/html");
?>
<style type="text/css">
<!--
.style2 {font-size: 14px}
-->
</style><head>
<title>Free Ringtones and HTML? And How did you get here?</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: large;
	color: #993300;
	font-weight: bold;
}
a {
	font-family: Arial, Helvetica, sans-serif;
	font-size: medium;
	color: #990000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #CCCCFF;
}
a:hover {
	text-decoration: none;
	color: #9999FF;
}
a:active {
	text-decoration: none;
	color: #0000FF;
}
-->
</style>

<?php 

    }
  } 

 if ($xmlyes == true) {
 ?>
<?php echo sprintf("%s%sxml version=\"1.0\"? encoding=\"utf\-8\"%s%s","<","?","?",">"); ?> 
<!DOCTYPE html PUBLIC "-//NOKIA//DTD XHTML Mobile +CHTML 1.0//EN" "http://www.nokia.com/dtd/xhtml-mp-chtml.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Ringtones? Want a Profile?</title>
<?php } ?>
</head>

<body>
<div align="center">
  <p><img src="../images/idents/ep_ident.jpg" width="101%" height="80%">
</p>
  <p class="style1"><a href="gprsdsl.php">Check for Broadband</a></p>
  <p class="style1"><a href="ringtones.php">Polyphonic Ringtones</a></p>
  <p class="style1 style2"><a href="wapartists.php?wapartist=1">WAP Artists</a></p>
  <p class="style1 style2">&nbsp;</p>
  <p class="style1">&nbsp;</p>
</div>
</body>
</html>
