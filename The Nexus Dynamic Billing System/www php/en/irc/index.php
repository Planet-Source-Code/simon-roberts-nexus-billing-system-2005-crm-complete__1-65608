<?php

	if ($nkey != "thickshake") {
		exit;
	}

	require_once('../../connections/projectalpha.php'); 

	mysql_select_db($database_projectalpha, $projectalpha);
	$query_sysop = sprintf("SELECT sysops.Username, sysops.Firstname, sysops.Surname FROM sysops WHERE sysops.RecID = %s",$SysopID);
	$sysop = mysql_query($query_sysop, $projectalpha) or die(mysql_error());
	$row_sysop = mysql_fetch_assoc($sysop);
	$totalRows_sysop = mysql_num_rows($sysop);
?>
<head>
<title>Dolphin Communications Fast Reliable Internet Solutions</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="Dolphin Communications, specialising in connecting the home or office user to the Internet on Fast and Reliable connections that do not cost the earth">
<meta name="keywords" content="web site design, web page design, custom web design, graphic design, adsl, flash, ecommerce,dialup, search, engine, development, search engine, web site, promotion, connection, internet connection">
<META name="revisit-after" content="15 days">
<META name="Robots" content="INDEX,FOLLOW">
<Meta name="distribution" content="global">

<LINK REL="SHORTCUT ICON" HREF="/favicon.ico">

<STYLE TYPE="text/css">
	<!--
	A {text-decoration: none; color: #00004a }
	A:hover {
	color:#999999;
}
td {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: normal; color: #00004A}
	-->
</STYLE>
<style type="text/css">
<!--
body {
{scrollbar-3d-light-color:#ACACAC;
scrollbar-arrow-color:#FFFFFF;
scrollbar-base-color:#ACACAC;
scrollbar-dark-shadow-color:#094882;
scrollbar-face-color:#243446;
scrollbar-highlight-color:#ACACAC;
scrollbar-shadow-color:#ACACAC;
}

.style1 {color: #00004A; font-size: 10px; font-weight: normal; }

.style2 {color: #00004a; font-family: Georgia, "Times New Roman", Times, serif;
	font-size: 14px;
	font-weight: bold; }
.style3 {color: #00004A; font-size: 10px; font-weight: bold; }
.style22 {font-size: 10px; color: #00004A;}
.style23 {color: #00004a}
.style25 {font-weight: normal; color: #00004A;}

-->

</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<noscript>Dolphin Communications specializes in creating solutions for all your internet and network design needs. Based in Sydney, Australia
we provide a very competitive price with a 1:16 contention rate giving you fast and reliable internet solutions. Complimented by our Dolphin Servers 
your needs are taken care of.</noscript>	

</head>

<body>
<applet code=IRCApplet.class archive="irc.jar,pixx.jar" width=980 height=728>
<param name="CABINETS" value="irc.cab,securedirc.cab,pixx.cab">

<param name="autoconnection" value="true">
<param name="nick" value="<?php echo $thisUsera; ?>">
<param name="alternatenick" value="<?php echo $thisUserb; ?>???">
<param name="name" value="<?php echo sprintf("%s %s",$row_sysop['Firstname'],$row_sysop['Surname']); ?>">
<param name="host" value="wollongong.oz.org">
<param name="gui" value="pixx">
<param name="port" value="6667">
<param name="language" value="language/english/index">
<param name="lngextension" value="php">
<param name="pixx:language" value="language/english/gui">
<param name="pixx:lngextension" value="php">

<param name="quitmessage" value="true">
<param name="asl" value="false">
<param name="useinfo" value="true">

<param name="style:bitmapsmileys" value="true">
<param name="style:smiley1" value=":) img/sourire.gif">
<param name="style:smiley2" value=":-) img/sourire.gif">
<param name="style:smiley3" value=":-D img/content.gif">
<param name="style:smiley4" value=":d img/content.gif">
<param name="style:smiley5" value=":-O img/OH-2.gif">
<param name="style:smiley6" value=":o img/OH-1.gif">
<param name="style:smiley7" value=":-P img/langue.gif">
<param name="style:smiley8" value=":p img/langue.gif">
<param name="style:smiley9" value=";-) img/clin-oeuil.gif">
<param name="style:smiley10" value=";) img/clin-oeuil.gif">
<param name="style:smiley11" value=":-( img/triste.gif">
<param name="style:smiley12" value=":( img/triste.gif">
<param name="style:smiley13" value=":-| img/OH-3.gif">
<param name="style:smiley14" value=":| img/OH-3.gif">
<param name="style:smiley15" value=":'( img/pleure.gif">
<param name="style:smiley16" value=":$ img/rouge.gif">
<param name="style:smiley17" value=":-$ img/rouge.gif">
<param name="style:smiley18" value="(H) img/cool.gif">
<param name="style:smiley19" value="(h) img/cool.gif">
<param name="style:smiley20" value=":-@ img/enerve1.gif">
<param name="style:smiley21" value=":@ img/enerve2.gif">
<param name="style:smiley22" value=":-S img/roll-eyes.gif">
<param name="style:smiley23" value=":s img/roll-eyes.gif">
<param name="style:backgroundimage" value="http://www.projectalpha.com.au/images/idents/ep_ident.jpg">
<param name="style:backgroundimage1" value="all all 0 background.gif">
<param name="style:sourcefontrule1" value="all all Serif 12">
<param name="style:floatingasl" value="true">

<param name="pixx:timestamp" value="true">
<param name="pixx:highlight" value="true">
<param name="pixx:highlightnick" value="true">
<param name="pixx:nickfield" value="true">
<param name="pixx:styleselector" value="true">
<param name="pixx:setfontonstyle" value="true">
<param name="pixx:showhelp" value="false">
<param name="pixx:showabout" value="false">
<param name="pixx:showchanlist" value="true">
<param name="pixx:showclose" value="false">
<param name="pixx:showstatus" value="true">

<?php
	if ($xoops_ircConfig['autojoin1']=='true') {
echo "<param name='command1' value='/join #".$channel1."'>";
	if ($xoops_ircConfig['autojoin2']=='true') {
echo "<param name='command2' value='/join #".$channel2."'>";
	} else {
}
	} else {
}
?>
</applet>
</body>
</html>

