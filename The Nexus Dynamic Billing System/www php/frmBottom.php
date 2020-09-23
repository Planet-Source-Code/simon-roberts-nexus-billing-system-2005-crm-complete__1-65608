
<?php

$hn = gethostbyaddr($REMOTE_ADDR);
if (stristr($hn, "no1.com.au")) {
Header("HTTP/1.1 404 Not Found");
print "<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML 2.0//EN\">
	<HTML><HEAD>
	<TITLE>404 Not Found</TITLE>
	</HEAD><BODY>
	<H1>Not Found</H1>
	The requested URL / was not found on this server, Or you do not have access or your domain is banned..<P>
	<P>Additionally, a 404 Not Found
	error was encountered while trying to use an ErrorDocument to handle the request.
	<HR>
	<ADDRESS>Apache/1.3.26 Server at <your_domain> Port 80</ADDRESS>
	</BODY></HTML>";
	exit;
}
 // Because this script sends out HTTP header information, the first characters in the file must be the <? PHP tag.

  $htmlredirect = "http://www.projectalpha.com.au/thenexus.php";                          // relative URL to your HTML file
  $wmlredirect = "http://www.projectalpha.com.au/wap/index.php";         // ABSOLUTE URL to your WML file

  if(strpos(strtoupper($HTTP_ACCEPT),"VND.WAP.WML") > 0) {        // Check whether the browser/gateway says it accepts WML.
    $br = "WML";
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
      $browser=="My S" ||                       // unknown Ericsson devkit browser ?
      $browser=="WAPJ" ||			// Virtual WAPJAG www.wapjag.de
      $browser=="fetc" ||			// fetchpage.cgi Perl script from www.wapcab.de
      $browser=="ALAV" ||			// yet another unknown UP based browser ?
      $browser=="Wapa")                         // another unknown browser (Web based "Wapalyzer"?)
        {
        $br = "WML";
    }
    else {
      $br = "HTML";
    }
  }
?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>


<title>The Nexus Gateway: Redirecting to PHP or WAP</title>
<HEAD>
<title>[<?php echo $row_Version['ver']; ?>] - The Nexus Project 2005</title>
<?php include('meta.php3'); ?>
<?php
  if($br == "WML") {
  	?>
<META HTTP-EQUIV="refresh" CONTENT="1;URL=<?php echo sprintf("%s",$wmlredirect); ?>">
</HEAD>
    <?php
  }
  else {
      	?>
<META HTTP-EQUIV="refresh" CONTENT="23;URL=<?php echo sprintf("%s",$htmlredirect); ?>">
</HEAD>
    <?php

  }
?>
<body>

<div align="center">
  <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#000099">
      <td bgcolor="#000099">&nbsp;</td>
      <td>&nbsp;</td>
      <td bgcolor="#000099">&nbsp;</td>
    </tr>
    <tr>
      <td height="990%" bgcolor="#000099">&nbsp;</td>
      <td><div align="center">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="96%" height="500">
          <param name="movie" value="introseq.swf">
          <param name=quality value=high>
          <embed src="introseq.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="96%" height="500"></embed>
        </object>
        <br>
          <a href="thenexus.php">(click to enter)</a> <br>
          <br>
          <SCRIPT LANGUAGE=javascript>
		  	setTimeout('document.all.cdown.innerHTML = 29;',1000);
			setTimeout('document.all.cdown.innerHTML = 28;',2000);
			setTimeout('document.all.cdown.innerHTML = 27;',3000);
			setTimeout('document.all.cdown.innerHTML = 26;',4000);
			setTimeout('document.all.cdown.innerHTML = 25;',5000);
			setTimeout('document.all.cdown.innerHTML = 24;',6000);
			setTimeout('document.all.cdown.innerHTML = 26;',7000);
			setTimeout('document.all.cdown.innerHTML = 25;',8000);
			setTimeout('document.all.cdown.innerHTML = 24;',9000);
			setTimeout('document.all.cdown.innerHTML = 23;',10000);
			setTimeout('document.all.cdown.innerHTML = 22;',11000);
			setTimeout('document.all.cdown.innerHTML = 21;',12000);
			setTimeout('document.all.cdown.innerHTML = 20;',13000);
			setTimeout('document.all.cdown.innerHTML = 19;',14000);
			setTimeout('document.all.cdown.innerHTML = 18;',15000);
			setTimeout('document.all.cdown.innerHTML = 17;',16000);
			setTimeout('document.all.cdown.innerHTML = 16;',17000);
			setTimeout('document.all.cdown.innerHTML = 15;',18000);
			setTimeout('document.all.cdown.innerHTML = 14;',19000);
			setTimeout('document.all.cdown.innerHTML = 16;',20000);
			setTimeout('document.all.cdown.innerHTML = 15;',21000);
			setTimeout('document.all.cdown.innerHTML = 14;',22000);
			setTimeout('document.all.cdown.innerHTML = 13;',23000);
			setTimeout('document.all.cdown.innerHTML = 12;',24000);
			setTimeout('document.all.cdown.innerHTML = 11;',25000);
			setTimeout('document.all.cdown.innerHTML = 10;',26000);
			setTimeout('document.all.cdown.innerHTML = 9;',27000);
			setTimeout('document.all.cdown.innerHTML = 8;',28000);
			setTimeout('document.all.cdown.innerHTML = 7;',29000);
			setTimeout('document.all.cdown.innerHTML = 6;',30000);
			setTimeout('document.all.cdown.innerHTML = 5;',31000);
			setTimeout('document.all.cdown.innerHTML = 4;',32000);
			setTimeout('document.all.cdown.innerHTML = 3;',33000);
			setTimeout('document.all.cdown.innerHTML = 2;',34000);
			setTimeout('document.all.cdown.innerHTML = 1;',35000);
			setTimeout('document.all.cdown.innerHTML = 0;',36000);
			setTimeout('window.location.href = "<?php echo sprintf("%s",$htmlredirect); ?>";',37000);
			 
			
			    </SCRIPT>
  <font face=arial size=2 color=black>This window will redirect <font color=red size=3><b><span id=cdown>30</span></b></font> seconds!</font></p></div></td>
      <td bgcolor="#000099">&nbsp;</td>
    </tr>
    <tr bgcolor="#000099">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
