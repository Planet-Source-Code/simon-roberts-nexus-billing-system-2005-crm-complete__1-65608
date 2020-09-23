<?php require_once('connections/projectalpha.php'); ?>
<?php require_once('Connections/Epwebdev.php'); ?>
<?php
mysql_select_db($database_projectalpha, $projectalpha);
$query_Version = "SELECT MAX(upgrade.Version) as Ver FROM upgrade";
$Version = mysql_query($query_Version, $projectalpha) or die(mysql_error());
$row_Version = mysql_fetch_assoc($Version);
$totalRows_Version = mysql_num_rows($Version);

header("Location: http://www.projectalpha.com.au/");
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

<LINK REL="SHORTCUT ICON" HREF="/favicon.ico">
<title>project alpha</title>
<meta http-equiv="CACHE" content="text/html; charset=iso-8859-1; PUBLIC">
<META NAME="GOOGLEBOT" CONTENT="ALL">
<META NAME="ROBOTS" CONTENT="ALL"> 
<META NAME="KEYWORDS" 
CONTENT="Exitstencil Press Pty Ltd, Dolphin Communications, Dolphin Soltuions, ADSL, SHDSL, SH-DSL, gSHDSL, Broadband, Telstra, XOOPS, WOOT, NO1, Comcen, TPG, VOIP, SIP, PABX, Mainframe, AS/400, zSeries, DMT, MDMA, Amphetamine, Phone Systems, Simon Roberts, Jarrett Costi, Brad Bulters, Andrew Riddock, The Nexus, Mobile, Telco, Telecom, AAPT, Vodophone, Telstra, Hudsonson, Satellite, T1, Maths, Science, i-ching, philopshy">
<META HTTP-EQUIV="EXPIRES" 
CONTENT="31 Jul 2005 24:00:00 GMT 10+">
<META NAME="DESCRIPTION" 
CONTENT="The Nexus 2004 - Highlighting designer plugins for new PNL and RFID Technologies. ">
<META NAME="COPYRIGHT" CONTENT="&copy; 1972 - 2004 - Exitstencil Press Pty Ltd - All Rights Reserved.">
<META NAME="AUTHOR" CONTENT="Simon A. Roberts, Sydney, Australia - admin@projectalpha.com.au">
<style type="text/css">
<!--
body,td,th {
	color: #0099FF;
}
body {
	background-color: #FFFFFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 01px;
	margin-bottom: 0px;
	background-image: url(images/backgrnd/<?php 
		$ntile = rand(1,26);
		echo sprintf("tile%d.jpg",$ntile); 
			$ttetextcolour = $ttetextcolour = "#FFCCCD";
			$subtextcolour = $subtextcolour = "#66DCFF";
		?>)
	
	
}
.style3 {
	color: #000099;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
.style4 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style5 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style7 {
	color: <?php echo $ttetextcolour; ?>;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	font-size: 12px;
}
.style8 {
	font-size: 48px;
	color: <?php echo $ttetextcolour; ?>;
}
.style11 {color: <?php echo $subtextcolour; ?>}
.style13 {
	font-size: 10px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #CCCCCC;
	font-weight: bold;
}
.style14 {font-family: Geneva, Arial, Helvetica, sans-serif; font-weight: bold; color: #FFFFFF; }
.style16 {
	font-size: 10;
	font-weight: bold;
}
.style18 {font-size: 10px; font-weight: bold; }
.style19 {
	font-size: 24px;
	color: #FFCC00;
}
.style22 {font-size: 24px; color: #000099; }
.style23 {
	color: #FF9933;
	font-weight: bold;
	font-size: 16px;
	font-style: italic;
}
.style24 {font-size: 24px}
.style27 {font-size: 18px}
.style28 {color: #663399}
-->
</style>


</head>

<body>
<table width="30%"  border="0" align="right" bgcolor="#FFFFFF">
  <tr>
    <td><div align="center"><a href="http://mail.projectalpha.com.au/"><img src="/images/icons/mail.jpg" width="49" height="41" border="0"> </a> <img src="images/whtspacer.GIF" width="10" height="33"><a href="en/login.php"><img src="/images/icons/cd-rom.jpg" width="38" height="38" border="0"></a><img src="images/whtspacer.GIF" width="10" height="33"> <a href="/projectalpha.php"><img src="/images/icons/projectalpha.jpg" width="44" height="40" border="0"></a> <img src="images/whtspacer.GIF" width="10" height="33"><a href="software.php"><img src="/images/icons/formulas.jpg" width="40" height="41" border="0"></a><img src="images/whtspacer.GIF" width="10" height="33">  <a href="/forum/" target="_self"><img src="/images/icons/contactus.jpg" width="38" height="40" border="0"></a> </div></td>
  </tr>
</table>
<table width="46%"  border="0" align="center">
  <tr>
    <td class="style4"><div align="center">
	<?php include("http://www.projectalpha.com.au/newsfeed.php?level=0&bckcolour=FFCC00"); ?>
</div></td>
  </tr>
  <tr>
    <td class="style4"><div align="center" class="style23">        </div></td>
  </tr>
</table>
<table width="76%"  border="0" align="center" bgcolor="#FFFFFF">
  <tr>
    <td><blockquote>
      <p align="justify" class="style3">&nbsp;</p>
      <p align="justify" class="style3"><strong><span class="style11"><span class="style23"><span class="style24">The Nexus Project</span> 2005<br>
        <br> 
        </span>Advanced Online Billing &amp; Moment of Sale Terminal </span><br>
          <br>
      </strong><span class="style7">For more information on viewing the<br>
inner workings of The Nexus, please view our <a href="pps/projectalpha.htm">slide show</a>.</span><br>
<br>
<span class="style8"> <span class="style7">[ release <?php echo $row_Version['Ver'];
	echo ' ]<br>';
	print realpath("http://www.projectalpha.com.au/index.php");
 ?></span></span> </p>
      <p class="style3">          <strong><em><br>
            <span class="style27">What is the Nexus? </span></em></strong></p>
      <p align="justify" class="style3"><strong><em>In the dictionary the word 'Nexus' is described as - A
          connected series or group. It was described by the Wall Street Journal as - A
          means of connection; a link or tie: &ldquo;this nexus between New York's... real-estate
investors and its... politicians&rdquo;. Bill Barol referred to the word 'Nexus' being -
            The core or center: &ldquo;The real nexus of the money culture Wall Street&rdquo;.</em></strong></p>
      <blockquote>
        <p align="justify" class="style3 style28"><em><a href="http://www.projectalpha.com.au/en/login.php">Log in and download</a> The Nexus 2005. If you have a business you can use The Nexus, once your business is setup in the system with it own identity and primary sysop account. Your business will be included in the <a href="http://www.projectalpha.com.au/biznet.php">Business Network</a> and product will be control and audited by the system. You can even use it for product placement. </em></p>
      </blockquote>
      <p align="justify" class="style3">Welcome to the new age in business an awakening of possibilities. As the world changes so does the technology. With introduction of RFID in warehouse and pick and pack stations to reduce shrinkage. Shortly the introduction of new internet technologies to be unveiled at technology shows in your cities will mean there will be little need for teir 1 in  End User Support of Digitally subscriber services. Such as the one catered in The Nexus because the internet will be told about your new Modem and be configured automatically when plugged in. </p>
      <p align="justify" class="style3">We have used <strong>Virtual World Networking</strong> (VWN) concept in the database design, so as a relational system goes it allow any type of service to be ordered on moment of sale, then subsequently billed accordantly if the item has a cycle state. This means that purchase orders, invoices, quota warnings, receipts and other such stationary are emailed to the client. The documents based on this stationary are live so they can be generated at any time and printed anytime from the <strong>Document Explorer</strong>.</p>
      <p align="justify" class="style3">You can also have your own private network of Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP, this means that your <strong>community</strong> of Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP's get access to all the features you are use to. What you charge them for the <strong>privileges</strong> is money you keep. The entire system is excluding Tax, which means that the simple double click action on any money field will bring up the TAX Calculator. This is done so you can easily trade with overseas clients.</p>
      <p align="justify" class="style3">Add your own <strong>Templates</strong> for other Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP's to resell from your business contacts allowing for a more dynamic and <em>larger scale business</em>, add your own <strong>vendors</strong>, add your own clients. All easily and fast. The <strong>MySQL</strong> back end ensure your data is <em>encrypted</em> using SSL between the The Nexus client and the server. This also means that the database can be queried by <strong>apache, q-mail, pureftp</strong> and all the endless titles of unix software available on the market. </p>
    </blockquote>      </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center">
      <p><img src="images/dialogues/MDIMainForm_1024x768.jpg" width="743" height="570"></p>
      <p>        <span class="style5">The Nexus's MDI Form</span></p>
    </div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><blockquote>
      <p align="justify" class="style3">Being a MySQL Client allows for SSL to be established over the Internet Protocol TCP/IP. This means that all your important and confidential data is kept that way. With a simple request we can enable your access to the templates screens in your Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP Accounts. This will mean that you can create your own Templates for the sales and billing procedural channels.</p>
      <p align="justify" class="style3">Once a template has been loaded into the sales channel for your group of sysop's, they will have access to these as an item for sale. This is then selectable as a service or subscription to use as a sales item.</p>
      <p align="justify" class="style3">Version 300 brings out a who new horizon, bug fee, hassle free, uncrashable, classed based error checking with database reference checking so and error will only occur once. Login by click on the CD-Rom and download today! </p>
    </blockquote></td>
  </tr>
  <tr>
    <td><div align="right"><br>
          <span class="style13"> Copyright 2001 - 2005 &copy; Exitstencil Press Pty Ltd <br>
        All Rights Reserved</span> </div></td>
  </tr>
</table>
<table width="76%" border="0" align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
  <tr bgcolor="#FFFFFF">
    <td>&nbsp;</td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="36%">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF" class="style8">
    <td colspan="3" class="style14"><div align="center"><span class="style22">Associates and Partnerships<br>
      <br>
    </span></div></td>
  </tr>
  <tr bgcolor="#FFFFFF" class="style8">
    <td colspan="3" class="style14"><div align="center" class="style19"></div></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td><div align="center"><span class="style18">    </span><span class="style16"><span class="style19"><img src="images/icons/top_left.gif" width="247" height="69"></span> </span>         </div></td>
    <td width="28%"><div align="center"><span class="style19"></span><br>
        </div></td>
    <td width="36%"><div align="center">
      <p><img src="images/idents/comcen_logo.JPG" width="159" height="137"></p>
      </div></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="3"><div align="center"><span class="style19"><img src="images/idents/westpac.bb.jpg" width="501" height="73"></span></div></td>
  </tr>
  <tr>
    <td height="97" colspan="3"><div align="center">
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </div></td>
  </tr>
</table>
<div align="center"><img src="images/galaxy_trans.gif" width="480" height="493">
</div>
</body>
</html>
<?php
mysql_free_result($Version);
?>
