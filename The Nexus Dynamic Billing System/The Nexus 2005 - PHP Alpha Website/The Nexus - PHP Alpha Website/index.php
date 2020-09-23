
<?php
error_reporting(0);
$hn = gethostbyaddr($REMOTE_ADDR);
if (stristr(" " . $hn, "no1.com.au")) {
Header("HTTP/1.1 404 Not Found");
print "<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML 2.0//EN\">
	<HTML><HEAD>
	<TITLE>404 Not Found</TITLE>
	</HEAD><BODY>
	<H1>Not Found</H1>
	Your Host ";
print $REMOTE_ADDR;
	print " has come up on our domain blacklist as one of the companies or network associate with bad debit to our network, This ban will be lifted when the work provided is settled legally, by our own auditorines. offending group, individual or company.";
	print "<br><br>The requested URL / was not found on this server, Or you do not have access or your domain is banned..<P>
	<P>Additionally, a 404 Not Found
	error was encountered while trying to use an ErrorDocument to handle the request.
	<HR>
	<ADDRESS>Apache/1.3.26 Server at <your_domain> Port 80</ADDRESS>
	</BODY></HTML>";
	exit;
}
 // Because this script sends out HTTP header information, the first characters in the file must be the <? PHP tag.

  $htmlredirect = "http://www.projectalpha.com.au/main.php";                          // relative URL to your HTML file
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

<?php 
if ($br == 'HTML') {
?>
<title>Simon Roberts - Why Pirate - Software Consultancy</title>
<?php } else { ?>
<title>The Nexus Project 2005</title>
<?php } ?>
<HEAD>

<?php
  if($br == "WML") {
  	?>
<META HTTP-EQUIV="refresh" CONTENT="1;URL=<?php echo sprintf("%s",$wmlredirect); ?>">
</HEAD>
    <?php
  }
  else {
      
// nexpv01 - Keywords Intial Tag Inclusion
// nexpv02 - Description Intial Tag Inclusion
include(sprintf("http://www.projectalpha.com.au/meta.php?nexpv01='%s'&nexpv02='%s'",'GPRS%20or%20HTML%20Redirector', 'GPRS%20or%20HTML%20Redirector')); 
?>
<META HTTP-EQUIV="refresh" CONTENT="23;URL=<?php echo sprintf("%s",$htmlredirect); ?>">
</HEAD>
    <?php

  }
?>
<body bgcolor="#000000">


        <?php
  if($br == "WML") {?>
          <img src="http://www.projectalpha.com.au/images/idents/ep_ident.jpg" width="99%" height="75%" border="0">
  <?php } else { ?>
  <div align="center">
  <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#000099">
      <td bgcolor="#000099">&nbsp;</td>
      <td>&nbsp;</td>
      <td bgcolor="#000099">&nbsp;</td>
    </tr>
    <tr>
      <td height="90%" bgcolor="#000099">&nbsp;</td>
      <td><div align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="62%" height="449">
          <param name="movie" value="http://www.projectalpha.com.au/introseq2.swf">
          <param name=quality value=high>
          <embed src="http://www.projectalpha.com.au/introseq2.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="62%" height="449"></embed>
        </object><?php } ?>
        <br>
        <font color="#009900"><strong><font size="4" face="Verdana, Arial, Helvetica, sans-serif">Consultancy &amp; Contracting</font><font size="4" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana, Arial, Helvetica, sans-serif"><br>
        Software Analayst
        </font></strong></font><font face="Verdana, Arial, Helvetica, sans-serif"><strong><br>
        <font color="#663366"><br>
        <font size="2">Simon Roberts, <br>
        Sydney, Australia</font></font> <br>
        </strong></font>
        <br>
          <a href="<?php
  if($br == "WML") {
         echo sprintf("%s",$wmlredirect);
  } else {
         echo sprintf("%s",$htmlredirect);
  } ?>">(click to enter)</a> <br>
          <?php
		include(sprintf('http://www.projectalpha.com.au/botredir.php?incdude=%s',gethostbyaddr($REMOTE_ADDR)));
	?>
          <br>
          <font color="#999999">
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
			setTimeout('window.location.href = "<?php
  if($br == "WML") {
         echo sprintf("%s",$wmlredirect);
  } else {
         echo sprintf("%s",$htmlredirect);
  } ?>";',37000);
			 
			
			    </SCRIPT>
          <font face=arial size=2>This window will redirect <font size=3><b><span id=cdown>30</span></b></font> seconds!</font>
          </p>
          </font></div></td>
        <?php if($br <> "WML") {?>
		<td bgcolor="#000099">&nbsp;</td>
    </tr>
    <tr bgcolor="#000099">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <?php } ?>
</div>
</body>
</html>
