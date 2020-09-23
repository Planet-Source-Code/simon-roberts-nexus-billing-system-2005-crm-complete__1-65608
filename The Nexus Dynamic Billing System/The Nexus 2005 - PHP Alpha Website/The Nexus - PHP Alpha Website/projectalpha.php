<?php require_once('connections/projectalpha.php'); ?>
<?php
mysql_select_db($database_projectalpha, $projectalpha);
$query_Version = "SELECT max(upgrade.Version) FROM upgrade";
$Version = mysql_query($query_Version, $projectalpha) or die(mysql_error());
$row_Version = mysql_fetch_assoc($Version);
$totalRows_Version = mysql_num_rows($Version);
?>
<<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>The Nexus Debreif</title>
<?php
// nexpv01 - Keywords Intial Tag Inclusion
// nexpv02 - Description Intial Tag Inclusion

$nexpv01 = 'The Nexus Debreif';
$nexpv02 = 'The Nexus Debreif';
include(sprintf("http://www.projectalpha.com.au/meta.php?nexpv01='%s'&nexpv02='%s%s'",str_replace(' ','%20',$nexpv01), str_replace(' ','%20',$nexpv02))); 
?>
<style type="text/css">
<!--
body,td,th {
	font-size: 14px;
	color: #000099;
}
body {
	background-color: #FFFFFF;
	margin-left: 2px;
	margin-top: 2px;
	margin-right: 10px;
	margin-bottom: 2px;
	background-image: url(images/backgrnd/tile8.jpg);
}
a:link {
	color: #0099FF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0000FF;
}
a:hover {
	text-decoration: none;
	color: #33CCCC;
}
a:active {
	text-decoration: none;
	color: #00CCCC;
}
.style12 {font-size: 18px}
.style14 {
	font-size: 12px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
.style17 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style18 {font-size: xx-large; font-family: Geneva, Arial, Helvetica, sans-serif; color: #FFFFFF; }
.style19 {font-size: large}
.style20 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style21 {font-size: large; font-family: Geneva, Arial, Helvetica, sans-serif; color: #FFCC99; }
.style22 {font-family: "Trebuchet MS", Tahoma, Arial}
.style27 {font-size: xx-large; font-family: Geneva, Arial, Helvetica, sans-serif; color: #FFCC99; }
.style28 {color: #FFCC99}
.style29 {font-size: 12px; font-family: Verdana, Arial, Helvetica, sans-serif; color: #FFCC99; }
.style30 {color: #FFFFFF}
.style31 {font-size: 12px; font-family: Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; }
-->
</style></head>

<body>
<table width="80%"  border="0" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td width="4%" rowspan="2">&nbsp;</td>
    <td colspan="5" rowspan="2"><p align="left" class="style27">The Nexus 2005 <br>
        <span class="style12">Billing Server and Moment of Sale Terminal</span></p></td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2"><div align="center" class="style28"><span class="style14"><strong>
      <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="410" height="284">
        <param name="movie" value="en/wizard2_option1.swf">
        <param name=quality value=high><param name="BGCOLOR" value="#000000">
        <embed src="en/wizard2_option1.swf" width="410" height="284" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#000000"></embed>
      </object>
      <br>
        <em><br>
    Captivating a full subscription inventory system, <strong>The Nexus 2005</strong>. It offers some of the most versitile billing system for the materials, telecommunications and consumables system. Linking credit card gateways and online browsing of the database. </em></strong></span></div></td>
    <td>&nbsp;</td>
    <td colspan="3"><p align="justify" class="style14 style30">&nbsp;</p>
      <p align="center" class="style18 style19"><strong>Imagine No more lost of data from crashes. </strong></p>
      <p align="left" class="style14 style30"><strong>::[ Highlights ]:: </strong></p>      
      <div align="justify">
        <ul>
          <li class="style31"><strong>Secure safe data, complete SSL between client and server. Complete confidental system. </strong></li>
          <li class="style31"><strong>Full Virtual World Complaint.</strong></li>
          <li class="style31"><strong>Customisable Vendors, Sales Templates, Sales Channels.</strong></li>
          <li class="style31"><strong>Document browser.</strong></li>
          <li class="style31"><strong>Hierarchal ViSP Network of your own in your own Virtual World.</strong></li>
          <li class="style31"><strong>Sysop Control Centre.</strong></li>
          <li class="style31"><strong>Invoice system with Reciepts, Statements.</strong></li>
          <li class="style31"><strong>Vendor System with complete XML Purchase orders.</strong></li>
          <li class="style31"><strong>Domain Information stored as a XML Pointer Document, this includes phone numbers, address and what ever other XML file you want to use with our &lt;XML&gt; Editor.</strong></li>
          <li class="style31"><strong>Complete MySQL server backend, set up anywhere on Unix&reg; or Microsoft Windows&reg;.</strong></li>
          <li class="style31"><strong>Full Expenditure System with linked Categories and Remittance Function.</strong></li>
          <li class="style31"><strong>Complete Prepayment, Client Deductable Bank Vault and accounts Recievables.</strong></li>
        </ul>
      </div>      <ul>
    </ul></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td class="style29">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td colspan="3"><div align="center"></div></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td colspan="2" class="style29"><p><strong>::[ Sysop Login ]:: </strong></p>
      <blockquote>
        <p align="justify" class="style14">The software has complete private networks of Sysops or administrators of the <strong>The Nexus 2005</strong> Virtual World you have access to. You can change ownership of Clients or Subscribers, between your sysops and ViSP's.</p>
        <p align="justify" class="style14">Control the level of access to the Virtual ISP network and the software from a percentage of information and control, to what form and sections of the software your sysops have. </p>
        <p align="justify" class="style14">Full network sweeps and Hot Link Encryption so everyting including your email addresses are encrypted in the database. 128bit password encryption or 256bit patchable password Encryption.</p>
        <p align="center"><strong>All over a SSL tcp/ip tunnel, so your information is kept confidential, private and safe from any packet sniffers or network monitor tools.</strong></p>
        <p align="justify" class="style14">Toolbar icon and short cut key is available for this form. This is so you can run the WAN terminal as a moment of sale terminal in a busy shop or office. This client tracks whos subscriber is whois.</p>
      </blockquote></td>
    <td>&nbsp;</td>
    <td colspan="3"><div align="center"><img src="images/dialogues/sysop_login.JPG" width="353" height="301"></div></td>
  </tr>
  <tr>
    <td colspan="7"><span class="style27"><br>
        <span class="style19">Commisson Calculator</span></span></td>
  </tr>
  <tr>
    <td colspan="7"><div align="center" class="style28">
      <p>&nbsp;</p>
      <p><img src="images/dialogues/commcentre.JPG" width="784" height="452"></p>
      <blockquote>
        <p align="left" class="style17">Thats right the commission calculator is a scales system for the re-accuring revenue generated by the subscriptions in the database. If you have access to the <strong>Rates and Scales </strong>frame. If your percentage of access you can set rates of commission per quanity billed. With live running financial information for Margin and percentiles of profit. Some Countries it is illegal to pay commission over a certain percentage. Australia is one of these.</p>
        <p align="justify" class="style14"><strong>Class based assignment to individual sysops.</strong> That right setup a commission matrix and assign to each sysop in your ViSP Network. Includes bonuses and spiff's for that keen sales driven business. Click on the dates you would like the report to be generated from then click <strong>Calculate from Date Selected.</strong> This will quickly scan the database and produce a drill down report on your sysops based on what Commission Class they get. You can then Export this Report as a CSV.</p>
        <p align="justify" class="style14">Microsoft Agent &reg; will direct you through the form and allow easy mouse over information about this form. </p>
        <p align="justify" class="style14">&nbsp;</p>
      </blockquote>
    </div></td>
  </tr>
  <tr>
    <td colspan="7"><span class="style21">Maintenance and Upkeep</span></td>
  </tr>
  <tr>
    <td colspan="7"><div align="center" class="style28">
      <blockquote>
        <p align="justify" class="style17">&nbsp;</p>
        <p align="justify" class="style20">The system is maintained by various drones or otherwise know as bot. From generating new invoices, to sending out mail. Some mail includes receipts for any money transaction in the system. Purchase order are also done at maintenance time, not everyone will have access to this system as it is more of a sub system. </p>
        </blockquote>
      <p><img src="images/dialogues/upkeep.JPG" width="623" height="227"></p>
    </div></td>
  </tr>
  <tr>
    <td colspan="7"><span class="style27"><br>
        <span class="style19">Virtual ISP (ViSP&reg;)</span></span></td>
  </tr>
  <tr>
    <td colspan="7"><blockquote class="style28">
      <p align="justify" class="style19"><br>
        <img src="images/dialogues/vw)VirtualISP.JPG" width="848" height="543"></p>
      <p align="justify" class="style19"><span class="style14">Have your own private virtual world of ViSP's. Each one like you having access to there own record for maintenance and updating. Allow them access to your regional server and maintain there business community in a virtual world of there own. Share templates on Sales Channels within your hierarchies. <br>
          <br>
Charge a barter value or complete joining fee to use the network. The Access you grant to the new ViSP is called a Primary Account. Slighly different from the normal sysop account, it can change the permissions of other sysops. These other sysops are within his/her's Virtual World or Business Community. </span> </p>
      <p class="style14"><strong>This new ViSP must have a registered ACN or ABN to trade within Australia. </strong>Other countires have law like this as well. If need be we can put your registered business number <strong>(RBN)</strong> in the system and also set taxing methods for sales within your country. <strong>The Nexus 2005 </strong>has a full global tax calculator built in.</p>
      <p class="style14"><strong><img src="images/dialogues/TAXCalx.JPG" width="196" height="103"><br>
        <br>
        * If you ever see a field for money this is the Excluding Tax price so we can easily trade with other countries. </strong>If you need to work out the Ex tax price in a money field to load up the Ex Tax Price Calculator by double clicking on the<strong> money value field.</strong></p>
    </blockquote>      
    </td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="891" colspan="7"><blockquote class="style28">
      <p align="justify" class="style22"><br>
        Welcome to the new age in business an awakening of possibilities. As the world changes so does the technology. With introduction of RFID in warehouse and pick and pack stations to reduce shrinkage. Shortly the introduction of new internet technologies to be unveiled at technology shows in your cities will mean there will be little need for teir 1 in End User Support of Digitally subscriber services. Such as the one catered in The Nexus because the internet will be told about your new Modem and be configured automatically when plugged in. </p>
      <p align="justify" class="style22">We have used Virtual World Networking (VWN) concept in the database design, so as a relational system goes it allow any type of service to be ordered on moment of sale, then subsequently billed accordantly if the item has a cycle state. This means that purchase orders, invoices, quota warnings, receipts and other such stationary are emailed to the client. The documents based on this stationary are live so they can be generated at any time and printed anytime from the Document Explorer . </p>
      <p align="justify" class="style22">You can also have your own private network of Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP, this means that your community of Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP's get access to all the features you are use to. What you charge them for the privileges is money you keep. The entire system is excluding Tax, which means that the simple double click action on any money field will bring up the TAX Calculator. This is done so you can easily trade with overseas clients. </p>
      <p align="justify" class="style22">Add your own Templates for other Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP's to resell from your business contacts allowing for a more dynamic and larger scale business , add your own vendors , add your own clients. All easily and fast. The MySQL back end ensure your data is encrypted using SSL between the The Nexus client and the server. This also means that the database can be queried by apache, q-mail, pureftp and all the endless titles of unix software available on the market. </p>
      <p align="justify" class="style22">Being a MySQL Client allows for SSL to be established over the Internet Protocol TCP/IP. This means that all your important and confidential data is kept that way. With a simple request we can enable your access to the templates screens in your Resellers, Wholesalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP Accounts. This will mean that you can create your own Templates for the sales and billing procedural channels. </p>
      <p align="justify" class="style22">Once a template has been loaded into the sales channel for your group of sysop's, they will have access to these as an item for sale. This is then selectable as a service or subscription to use as a sales item. </p>
      <p align="center" class="style22">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="460" height="317">
          <param name="movie" value="en/wizard2_option2.swf">
          <param name=quality value=high><param name="BGCOLOR" value="#000000">
          <embed src="en/wizard2_option2.swf" width="460" height="317" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#000000"></embed>
        </object>
      </p>
    </blockquote></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="8%">&nbsp;</td>
    <td width="40%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
    <td width="9%">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<?php
mysql_free_result($Version);
?>
