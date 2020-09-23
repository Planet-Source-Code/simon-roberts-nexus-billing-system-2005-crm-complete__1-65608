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
<title>Software Development - Simon Roberts, Australia - IT Consultant &amp; Business Analyst</title>
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
.style9 {color: #CCCCCC}
.style11 {
	font-size: xx-large;
	color: #CCCCCC;
}
.style12 {font-size: medium}
-->
</style><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="29%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
</tr>
<table width="770" height="857" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999" bgcolor="#FFFFFF">
      <tr>
        <td height="44" colspan="2" bgcolor="#000000"><p align="right" class="style11">Software Consultancy </p></td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#FFFFFF">
  <tr>
            <td height="626" colspan="2" bgcolor="#020615" ><blockquote class="style12">
              <p align="justify">&nbsp;                </p>
              <p align="justify"> Have you or are you wanting to develop applications, have in-house analyses done on an application or looking to hire for a contract. Well you have come to the right place, I am here to hire. I am a Software Developer, specialising in secure and complex systems. We work in-house or in your office. Perusing our time within both compilers and other compiler utilities. </p>
              <p align="justify"> I work in VB 6.0, C#, VB.Net and have a range of tools and utilities that I develop along side. I also work with a variety of databases servers and stand-alone's like access. Some of the database services I am confluent with are Microsoft SQL Server, My SQL, Oracle 8i &amp; 9 as well as bTreve, Fox Pro and a host of other small time usages of the lesser mentioned utilities. </p>
              <p align="justify"> I can design web sites, in XML, HTML, ASP, ASP.Net, PHP as well as being resoundably useful when it comes to graphic design of the web site and correlating images with a Graphic Design Artists. </p>
              <p align="justify"><img src="/images/clip_image002.jpg" width="268" height="216" hspace="20" vspace="20" align="right"> Do you want some time dedicated to reporting, well that is no worries you can have me take a percentage of space to developing reports and justifications and terms for expenditure or time values commenting. </p>
              <p align="justify">I have been programming since the early days of computing starting on a commodore vic 20 as well as being extremely, verse in networking confluences and strata. I enjoy services to your business or home soho, with concise rudimentary changes to the maintenance's or your first application your project alpha or beta's of appliance and software. </p>
              <p align="justify">I am currently living in sydney, wanting to start a degree in Law at one of the universities in the sydney, this is so I can specialise in White Collar Crime as well as Computing Litigation. But this will be over 5 years of study, which i am looking forward to.</p>
              <p align="justify">I have worked both as a Business Analyst as well as Programming Analyst in various companies. Not to mention other lesser roles in business, like support analyst, technician, network engineer as well as a sales role here amongst the references.</p>
              <p align="justify">Developing a host of applications for you or your business is where I specialise in software development. I look forward to hearing from you and planning a viable resoundable solution that is confluent with your working practiced and computer usages.</p>
              <p align="justify">&nbsp;</p>
            </blockquote>            </td>
  </tr>
          <tr>
            <td height="35" colspan="2" bgcolor="#020615"><blockquote>
              <div align="center">              </div>
            </blockquote></td>
  </tr>
          <tr>
            <td width="85%" height="53" bgcolor="#020615"><ul>
              <li>Resume For IT Consultant- <a href="resumes/i_sar_resume_05wint.doc">Word 2000</a>, <a href="resumes/sar_resume_it_full.pdf">PDF Document</a>, <a href="resumes/it_sar_full.htm">HTML</a>.</li>
              <li>Resume For IT Consultant Abbreviation - <a href="resumes/i_sar_resume_short.doc">Word 2000</a>, <a href="resumes/sar_resume_abbv.pdf">PDF Document</a>, <a href="resumes/it_sar_short.htm">HTML</a>.</li>
              <li>Resume For Retail Support &amp; Wholesale  - <a href="resumes/i_sar_resume_retail.doc">Word 2000</a>, <a href="resumes/sar_resume_retail.pdf">PDF Document</a>, <a href="resumes/it_sar_retail.htm">HTML</a>.</li>
              <li>Resume For Radio - Word 2000, PDF Document, HTML.</li>
            </ul></td>
            <td width="15%" height="53" colspan="-1" align="center" bgcolor="#020615">              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="103" height="24">
              <param name="movie" value="_mail.swf">
              <param name="quality" value="high"><param name="BGCOLOR" value="#020615">
              <embed src="_mail.swf" width="103" height="24" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#020615" ></embed>
            </object>
              <br>
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="103" height="24">
                <param name="movie" value="button1.swf">
                <param name="quality" value="high">
                <param name="bgcolor" value="#020615">
                <embed src="button1.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="103" height="24" bgcolor="#020615"></embed>
              </object>
              <br>
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="103" height="24">
                <param name="movie" value="_software.swf">
                <param name="quality" value="high">
                <param name="bgcolor" value="#020615">
                <embed src="_software.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="103" height="24" bgcolor="#020615"></embed>
              </object>
              <br>
            </td>
          </tr>
          <tr>
            <td height="47" colspan="2" bgcolor="#020615"><blockquote>
              <p><span class="style9">Copyright 2001 - 2006 <strong>&copy;</strong> Why Pirate<br>
    All Rights Reserved, Intellectual property on content and psuedocode. </span>
                <?php
		include(sprintf('http://www.projectalpha.com.au/botredir.php?incdude=%s',gethostbyaddr($REMOTE_ADDR)));
	?>
              </p>
            </blockquote></td>
          </tr>
</table>
<p>&nbsp;</p>
        <p>&nbsp;</p>
</body>
</html>
