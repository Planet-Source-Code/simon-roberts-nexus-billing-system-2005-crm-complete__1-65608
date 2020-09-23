<?php require_once('connections/projectalpha.php'); ?>
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
	print " has come up on our domain blacklist as one of the companies or network associate with bad debit to our network, This ban will be lifted when the until legal persuit has been finalised. by the offending group, individual or company.";
	print "<br><br>The requested URL / was not found on this server, Or you do not have access or your domain is banned..<P>
	<P>Additionally, a 404 Not Found
	error was encountered while trying to use an ErrorDocument to handle the request.
	<HR>
	<ADDRESS>Apache/1.3.26 Server at <your_domain> Port 80</ADDRESS>
	</BODY></HTML>";
	exit;
}

mysql_select_db($database_projectalpha, $projectalpha);
$query_Version = "SELECT max(upgrade.Version) as ver FROM upgrade ";
$Version = mysql_query($query_Version, $projectalpha) or die(mysql_error());
$row_Version = mysql_fetch_assoc($Version);
$totalRows_Version = mysql_num_rows($Version);
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<title>[<?php echo $row_Version['ver']; ?>] - The Nexus Project 2005 ( Business Resource Sharing, Moment of Sale Terminal, POS)</title>
<?php
// nexpv01 - Keywords Intial Tag Inclusion
// nexpv02 - Description Intial Tag Inclusion

$nexpv01 = 'The Nexus 2005 Homepage';
$nexpv02 = 'The Nexus 2005 Homepage';
include(sprintf("http://www.projectalpha.com.au/meta.php?nexpv01='%s'&nexpv02='%s'",str_replace(' ','%20',$nexpv01), str_replace(' ','%20',$nexpv02))); 
?>

<style type="text/css">
<!--
body,td,th {
	font-family: Trebuchet MS, Tahoma, Arial;
	font-size: 12px;
	color: #FEAC89;
}
body {
	background-color: #000033;
	margin-left: 6px;
	margin-top: 3px;
	margin-right: 6px;
	margin-bottom: 0px;
	background-image: url(images/tile77d.jpg);
}
.style2 {font-size: 7px}
.style3 {
	font-size: 18px;
	color: #A2CDFD;
	font-weight: bold;
}
.style5 {
	font-size: small;
	font-weight: bold;
}
a:link {
	color: #FF3300;
	font-weight: bold;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #990033;
}
a:hover {
	text-decoration: none;
	color: #339933;
}
a:active {
	text-decoration: none;
	color: #000066;
}
.style6 {
	font-size: small;
	color: #FF3300;
}
.style7 {color: #CC3333}
.style8 {font-size: small}
.style9 {color: #CCCCCC}
-->
</style><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>

<body>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="38%"><img src="images/logo1.gif" width="347" height="110"></td>
    <td width="62%"><div align="center">
        <?php include("http://www.projectalpha.com.au/newsfeed.php?nVirtualID=0&level=0&bckcolour=FFCC00"); ?>
    </div></td>
  </tr>
  <tr>
    <td height="44" colspan="2"><div align="center"><span class="style2"></span>
            <table width="313" border="0" align="left" cellpadding="0" cellspacing="0">
              <tr>
                <td width="43" align="center" valign="middle"><a href="en/login.php"><img src="images/icons/Building_House_1.gif" width="32" height="32" border="0"></a> </td>
                <td width="38" align="center" valign="middle"><a href="http://mail.projectalpha.com.au/"><img src="images/icons/download.gif" width="32" height="32" border="0"></a></td>
                <td width="44" align="center" valign="middle"><a href="software.php"><img src="images/icons/mrblackboard.gif" width="32" height="32" border="0"></a></td>
                <td width="50" align="center" valign="middle"><a href="biznet.php"><img src="images/icons/contactus.jpg" width="32" height="32" border="0"></a></td>
                <td width="46" align="center" valign="middle"><a href="software.php"><img src="images/icons/softlab_32x32.gif" width="32" height="37" border="0"></a></td>
                <td width="46" align="center" valign="middle"><a href="DSLAvailable.php"><img src="images/icons/BIPAC-5100.gif" width="33" height="32" border="0"></a></td>
                <td width="46" align="center" valign="middle"><a href="http://www.projectalpha.com.au/forum/modules/newbb/"><img src="images/icons/dbtype.gif" width="32" height="32" border="0"></a></td>
              </tr>
            </table>
      </div>
        <div align="left"></div></td>
  </tr>
</table>
</td>
</tr>
<tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="29%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
</tr>
<table width="770" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#990000">
          <tr bgcolor="#00000D" >
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="login.swf">
              <param name="quality" value="high">
              <param name="bgcolor" value="#FFFFFF">
              <embed src="login.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF"></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="mail.swf">
              <param name="quality" value="high">
              <param name="bgcolor" value="#FFFFFF">
              <embed src="mail.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF"></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="BGCOLOR" value="">
              <param name="movie" value="button3.swf">
              <param name="quality" value="high">
              <embed src="button3.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" ></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="othersoftware.swf">
              <param name="quality" value="high">
              <param name="bgcolor" value="#FFFFFF">
              <embed src="othersoftware.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF"></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="dsl.swf">
              <param name="quality" value="high">
              <param name="bgcolor" value="#FFFFFF">
              <embed src="dsl.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF"></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="slideshow.swf">
              <param name="quality" value="high">
              <embed src="slideshow.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" ></embed>
            </object></td>
            <td height="22" align="center" valign="middle"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="22">
              <param name="movie" value="debrief.swf">
              <param name="quality" value="high">
              <embed src="debrief.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" ></embed>
            </object></td>
  </tr>
</table> 
		<table width="770" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999" bgcolor="#FFFFFF">
      <tr>
        <td width="808" height="44" bgcolor="#000000"><p align="center"><span class="style3">The Nexus Project 2005 - Business Networks for a New Tomorrow</span></p></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFFF">
        <tr>
            <td bgcolor="#020615" ><blockquote>
              <p align="left" class="style8"><br>[ release <?php echo $row_Version['ver']; ?> ] </p>
              <p class="style5">What is the Nexus? </p>
              <p align="justify" class="style5">In the dictionary the word 'Nexus' is described as - A connected series or group. It was described by the Wall Street Journal as - A means of connection; a link or tie: “this nexus between New York 's... real-estate investors and it's... politicians”. Bill Barol referred to the word 'Nexus' being - The core or center: “The real nexus of the money culture Wall Street”. </p>
              <blockquote class="style8">
                <p align="justify"><em><a href="http://www.projectalpha.com.au/en/login.php">Log in and download </a> <strong>The Nexus 2005</strong>. If you have a business you can use The Nexus, once your business is setup in the system with it own identity and primary sysop account. Your business will be included in the <a href="http://www.projectalpha.com.au/biznet.php">Business Network </a> and product will be control and audited by the system. You can even use it for product placement. </em></p>
              </blockquote>
              <table width="100%"  border="0" align="center" cellpadding="11" cellspacing="0">
                <tr valign="top">
                  <td width="50%" class="style8"><p><strong>What is The Nexus?</strong></p>
                    <div align="justify">
                      <p>The nexus is a new business tool for billing subscription or services. It can even be used to run the daily grind of a retail environment an ISP or another type of trade all together that has fixed pricing over periods of time.</p>
                      <p>This is a complete Software Solution to your reseller, billing and point of sale solution. Whether your running a call center/phone orders. A retail environment or a shipping operation then we can provide solutions for RFID or PNL systems being introduced to these industries soon.</p>
                      <p>Complete with scan and pack sequencing and back end maintenance that emails your invoices, purchase orders, receipts and other stationary that is crucial to any billing system.</p>
                      <p><strong>What is MOST Terminal?</strong></p>
                      <p><strong>Moment of Sale  Terminal (MOST)</strong> allows you and your associate system operators to do sales transactions, refunds, banking, reseller invoicing, your invoices, BNP Creation and a mirage of other function that are performed for sales of items.</p>
                      <p align="center"><img src="images/dialogues/MDIMainForm_350x280.jpg" width="268" height="216"><br>
                      <strong>MOST MDI Form </strong></p>
                      <p align="justify">MOST allow anyone with access to your information, this is a sysop that you have given permission to view levels of detail in the software. This will allow you to track your client and there associated transaction data. Includes full audit system.</p>
                      <p align="justify">&nbsp; </p>
                    </div>                    </td>
                  <td width="50%" class="style8"><p><strong>How do I get started? </strong></p>
                  <p align="justify">The first step of setting up your business and having it listed on the Business Network Providers (BNP) manifest is to register yourself an account on the <a href="en/login.php">Sysop Control Center</a>.</p>
                  <p align="justify">Once you have a logon you can either join a BNP or set up a new business in The Nexus which will allow you to access your own private tier in the platform. You can share some of your resources for other companies to sell on a 'referral' system or import a plan template from another company on the network.</p>
                  <p align="justify">You must have a Register Business Number in your country to set up a Business Network Provider. This provides your own private news feed and access to our software. </p>
                  <p align="justify"><strong>What is The Sales Channel Editor?</strong></p>
                  <p align="justify">The Nexus' Sales Channel Editor allow you as a primary sysop to maintain your products and get reports on performance of your network and sales information. </p>
                  <p align="justify">From here you can create your Services Product Tree from which you have your templates and sales items catalogued under.</p>
                  <p align="justify">This is a management tool for reporting and control of the sales MOST terminal and the products that it support. This is also where you can do Product Template Sharing where other companies can explore your services tree and select from product that you have source from your private vendors.</p>
                  <p align="center"><img src="images/dialogues/SCE_MainMDI.jpg" width="268" height="216"></p>
                  <p align="center"><strong>You will need to download the alpha of NSCE to build your sales tree and service tree. </strong></p></td>
                </tr>
                <tr>
                  <td align="center" valign="top" class="style5 style6"><p>                      <img src="en/images/warning.gif" width="15" height="15"> Upgrade Notice <img src="en/images/warning.gif" width="15" height="14"></p>
                  <p><span class="style7">You will need to upgrade due to security enhancements</span>.</p></td>
                  <td class="style8"><div align="justify">
                    <p>The Nexus Sales Channel Editor now support creation of VOIP Billing Templates/Profiles. These can interact with the sales terminal but on a lower scale until we update from the latest version. </p>
                    <p align="center"><strong>Try Version 2.0.0.5 today!</strong></p>
                  </div></td>
                </tr>
              </table>
              <p align="left" class="style8">So we have you somewhat interested well please <a href="thenexusmore.php">click here</a> to read more about the nexus and the potential to expand your business new or old to new horizons in business and commerce.</p>
          </blockquote>            </td>
          </tr>
          <tr>
            <td bgcolor="#020615"><blockquote>
              <div align="center"></div>
            </blockquote></td>
          </tr>
          <tr>
            <td bgcolor="#020615"><blockquote>
              
</blockquote></td>
          </tr>
          <tr>
            <td height="235" bgcolor="#020615"><blockquote>
              <p align="center"><span class="style9"><img src="images/galaxy_trans.gif" width="663" height="156"></span></p>
              <p align="left"><span class="style9">Copyright 2001 - 2005 &copy; Why Pirate<br>
        All Rights Reserved, Intellectual property on content and psuedocode. </span>
                <?php
		include(sprintf('http://www.projectalpha.com.au/botredir.php?incdude=%s',gethostbyaddr($REMOTE_ADDR)));
	?>
                <br>
              </p>
            </blockquote></td>
          </tr>
</table>
<p>&nbsp;</p>
        <p>&nbsp;</p>
</body>
</html>
