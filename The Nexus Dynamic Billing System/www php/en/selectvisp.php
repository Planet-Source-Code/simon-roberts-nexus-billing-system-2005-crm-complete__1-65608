<?php require_once('../Connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}

$mailsent = false;
require("CustomSql.inc.h4x0r.php");

if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">$front_pleaselogin</a>";
	exit;
} else {
	if ($visptochoose==0) {
	
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_rsSysopDetails = sprintf("select INClause, Username, md5(Decode(`Password`,'dr34mt1me')) as md5Pass from sysops where RecID = '%d' ", $_SESSION['SysopID']);
			$rsSysopDetails = mysql_query($query_rsSysopDetails, $projectalpha) or die(mysql_error());
			$row_rsSysopDetails = mysql_fetch_assoc($rsSysopDetails);
			$totalRows_rsSysopDetails = mysql_num_rows($rsSysopDetails);
			if (!($_SESSION['Reseller'] <> 0) || !session_is_registered('Reseller')){
				$_SESSION['Reseller'] = "yuk";
			}
		} else {
			
			$_SESSION['Reseller']=$visptochoose;
		}
	if ($_SESSION['Reseller'] <> 0) {
		
		if (!session_is_registered('Reseller') || $_SESSION['Reseller']=="yuk") {
			if (!($visptochoose==0)) {
				$_SESSION['Reseller']=$visptochoose;
			}
		}
		if (!empty($chksendnow)) {
		
		require("classes.php");
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsSysopDetails = sprintf("select Firstname, Surname, Email, Username, md5(Decode(`Password`,'dr34mt1me')) as md5Pass from sysops where RecID = '%d' ", $_SESSION['SysopID']);
		$rsSysopDetails = mysql_query($query_rsSysopDetails, $projectalpha) or die(mysql_error());
		$row_rsSysopDetails = mysql_fetch_assoc($rsSysopDetails);
		$totalRows_rsSysopDetails = mysql_num_rows($rsSysopDetails);
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsSysopDetails1 = sprintf("select virtualisp.Description, virtualisp.ABN, `sysops`.`RecID` as YSID, md5(Decode(`sysops`.`Password`,'dr34mt1me')) as md5Pass, sysops.Firstname, sysops.Surname, sysops.Email, sysops.Mobile from sysops inner join virtualisp on virtualisp.SysopID = sysops.RecID where virtualisp.RecID = '%d' ", $_SESSION['Reseller']);
		$rsSysopDetails1 = mysql_query($query_rsSysopDetails1, $projectalpha) or die(mysql_error());
		$row_rsSysopDetails1 = mysql_fetch_assoc($rsSysopDetails1);
		$totalRows_rsSysopDetails1 = mysql_num_rows($rsSysopDetails1);

		$subject = sprintf("Activation of Sales Channel For - [%s] - %s, %s",$row_rsSysopDetails['Username'],$row_rsSysopDetails['Surname'],$row_rsSysopDetails['Firstname']);
		$fullname = sprintf('%s %s - %s',$row_rsSysopDetails['Firstname'],$row_rsSysopDetails['Surname'],$row_rsSysopDetails['Username']);
		$editlink = sprintf('http://www.projectalpha.com.au/en/editperm.php?Email=%s&nYSID=%d&nMD5A=%s&nSysopID=%d',$row_rsSysopDetails1['Email'], $row_rsSysopDetails1['YSID'],$row_rsSysopDetails1['md5Pass'],$_SESSION['SysopID']);
		
		$emailbody = $emailmsg . chr(10) . chr(13) . "To Activate or Discard this application please click on this link: " . $editlink ;
		
		$xoopsMailer = & getMailer();
		$xoopsMailer->useMail();
		$xoopsMailer->setToEmails($row_rsSysopDetails1['Email']);
		$xoopsMailer->setFromEmail($row_rsSysopDetails['Email']);
		$xoopsMailer->setFromName($fullname);
		$xoopsMailer->setSubject($subject);
		$xoopsMailer->setBody($emailbody);
		$xoopsMailer->send();
		if ( !$xoopsMailer->send() ) {
			echo $xoopsMailer->getErrors();
		} else {
			echo $emailmsg . chr(10) . chr(13) . "To Activate or Discard this application please click on this link: ";
		}
		$mailsent=true;
	} else {
		
	}

}

?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>The Nexus Registration for Sysop and Bonus Rankings in Forum</title>
<meta http-equiv="Content-Type" content="text/html; charset=<?php print "$front_charset"; ?>">
<link rel="stylesheet" href="./style/style.css" type="text/css">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
<link href="../css/txtbox3.css" rel="stylesheet" type="text/css">
<link href="txtbox2.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style3 {font-weight: bold}
.style4 {color: #FF0000}
.style32 {color: #FFFFFF}
.style33 {
	font-size: 16px;
	font-weight: bold;
}
.style35 {font-size: 16px}
.style36 {color: #000000}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<div align="center">
  <?php
include("top.php3");
?>
  <table width="770"  border="0" align="center" cellpadding="0" cellspacing="0">
    <?php
	if ($_SESSION['Reseller']=="yuk") {
	?>
	<tr align="left" valign="top">
      <td colspan="2" bgcolor="#000000" class="txtbox style32"><div align="left">Select Your Sponsor </div></td>
    </tr>
    <tr align="left" valign="top" bgcolor="#FFFFFF">
      <td colspan="2"><blockquote>
        <p><br>
          This is the company that will be providing the resources for your sales items in the nexus. They will be your team and allies with The Nexus and ensure that you have the very best of this companies products to sell and generate a commission for. After you have selected the company that you want to be a sysop for then you will be required to send a brief cover letter to them introducing yourself, your interest and skill set. As they may be able to offer you employment or training.<br>
          <br>
          <span class="style3">Select from the following companies which one you wish to join:</span>
          </p>
        </blockquote></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td colspan="2"><form name="form1" method="post" action="">
          <div align="center">
            <p>
              <select  name="visptochoose" size="20" class="txtbox" id="visptochoose">
			  <?php
			  mysql_select_db($database_projectalpha, $projectalpha);
			$query_rsVISPs = sprintf("select distinct RecID in %s, Description, RecID from virtualisp",$_SESSION['INClause']);
			echo $query_rsVISPs;
			$rsVISPs = mysql_query($query_rsVISPs, $projectalpha) or die(mysql_error());
			$totalRows_rsVISPs = mysql_num_rows($rsVISPs);
			
			 while ($row_rsVISPs = mysql_fetch_assoc($rsVISPs)) {
			  ?>
                <option value="<?php echo $row_rsVISPs['RecID']; ?>"><?php echo $row_rsVISPs['Description']; ?></option>
				<?php } 
					
					mysql_free_result($rsSysopDetails);
					mysql_free_result($rsVISPs);
				?>
              </select>
            </p>
            <p align="center">
              <input name="Submit" type="submit" class="txtbox" value="Choose This Business Provider --&gt;">    
            </p>
          </div>
      </form></td>
    </tr>
	<?php
	} else {
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsSysopDetails = sprintf("select virtualisp.Description, virtualisp.ABN, sysops.Firstname, sysops.Surname, sysops.Email, sysops.Mobile from sysops inner join virtualisp on virtualisp.SysopID = sysops.RecID where virtualisp.RecID = '%d' ", $visptochoose);
		$rsSysopDetails = mysql_query($query_rsSysopDetails, $projectalpha) or die(mysql_error());
		$row_rsSysopDetails = mysql_fetch_assoc($rsSysopDetails);
		$totalRows_rsSysopDetails = mysql_num_rows($rsSysopDetails);
			
		if (!$mailsent) {
		
	?>
    <tr bgcolor="#000000" class="txtbox">
      <td colspan="2"><span class="txtbox style32">Your Sysop Reseller Joining Letter </span></td>
    </tr>
    <tr align="center" valign="middle" bgcolor="#CCCCFF">
      <td height="132" colspan="2"><blockquote>      
        <div align="justify">Well this is much like joining a company or a business through the job interview process. Here is where you introduce yourself to the company who you are wanting to be an associated sysop with. This letter will go to a senior management representive on the Nexus' Databases. This is much like the cover letter you send with a resume. If this person accepts your application to join there team on the Nexus then you will be paid commission and even possibly salary from this company this is all for you to discuss with them.<br>
            <br>
            <strong>Please make a recording of the primary sysop Name, Email Address and Mobile Number so you can chase up or return there call or contact them to thank you for accepting you to the Nexus on there Reseller Node. 
            </strong> 
        </div>
      </blockquote></td>
    </tr>
    <tr align="center" valign="top" bgcolor="#CCCCFF">
      <td width="99"><div align="right" class="style33">
        <h3>To:</h3>
      </div></td>
      <td width="671"><blockquote><div align="right" class="style33">
        <h3 align="left"><?php echo sprintf('%s %s',$row_rsSysopDetails['Firstname'],$row_rsSysopDetails['Surname']); ?></h3>
      </div></blockquote></td>
    </tr>
    <tr align="center" valign="top" bgcolor="#CCCCFF">
      <td><div align="right" class="style35">
        <h3><strong>Company:</strong></h3>
      </div></td>
      <td><blockquote><div align="right" class="style33">
        <h3 align="left"><?php echo sprintf('%s - %s',$row_rsSysopDetails['Description'],$row_rsSysopDetails['ABN']); ?></h3>
      </div></blockquote></td>
    </tr>
    <tr align="center" valign="top" bgcolor="#CCCCFF">
      <td><div align="right" class="style35">
        <h3><strong>Email:</strong></h3>
      </div></td>
      <td><blockquote><div align="right" class="style33">
        <h3 align="left"><?php echo $row_rsSysopDetails['Email']; ?></h3>
      </div></blockquote></td>
    </tr>
    <tr align="center" valign="top" bgcolor="#CCCCFF">
      <td><div align="right" class="style35">
        <h3><strong>Mobile:</strong></h3>
      </div></td>
      <td><blockquote><div align="right" class="style33">
        <h3 align="left"><?php echo $row_rsSysopDetails['Mobile']; ?></h3>
      </div></blockquote></td>
    </tr>
    <tr align="center" valign="top" bgcolor="#CCCCFF">
      <td><div align="right" class="style35"><strong>Message:</strong></div></td>
      <td><blockquote>
        <form name="form2" method="post" action="">
          <div align="right">
            <p align="center">
              <textarea name="emailmsg" cols="60" rows="40" wrap="VIRTUAL" class="txtbox" id="emailmsg">Dear <?php echo sprintf('%s %s',$row_rsSysopDetails['Firstname'],$row_rsSysopDetails['Surname']); ?>,

I was so excited to find you available for joining as part of your sales staff and i hope to establish through your company a firm and secure line of commission for the sales of your products by myself.

With joining your company as part of it personell I hope to achieve high sales standards and margin. Some of my technical or other skills are as follows that i would like to make available for paid work for your company:
  
  •
  •
  •
  •
  •

I understand that I may have to sign privacy statements and adhear to the privacy laws of Australia. I also understand that you may be required to formally interview me in person and get me involved in closer things within the company.

I have worked in the following places i have worked and you may contact as points of reference are:

  •    00/00/0000    -
  •
  •
  •
  •

I will wait to hear from you either by email or phone to notify myself of whether i have been given access to the sales channel available to your sysops on The Nexus 2005. If I do not hear from you within 14 days I will contact you by phone on <?php echo $row_rsSysopDetails['Mobile']; ?>.

Kind Regards


Your Name</textarea>
            </p>
            <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="59%" valign="middle"><strong>
                  <input name="chksendnow" type="checkbox" id="chksendnow" value="send">
                  I have checked my email and are ready to send </strong></td>
                <td width="41%"><div align="center">
                  <input type="submit" name="Submit2" value="Send Acceptance Email">
                </div></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <p align="left"><br>
              </p>
          </div>
        </form>
      </blockquote></td>
    </tr>
    <tr bgcolor="#CCCCFF">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<?php 
		
		} else {
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsSysopDetails1 = sprintf("update sysops set bConfirmed = '-2', VirtualID = '%d' where RecID = '%d' ", $_SESSION['Reseller'], $_SESSION['SysopID']);
		$rsSysopDetails1 = mysql_query($query_rsSysopDetails1, $projectalpha) or die(mysql_error());
		$_SESSION['bConfirmed']=-2;
		
		 ?>
    <tr bgcolor="#FFCC33">
      <td colspan="2" bgcolor="#000000" class="txtbox style32">Your email is sent! </td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td colspan="2"><blockquote>
        <p class="style36"><strong>Now all you have to do is wait until the primary sysop or his representive has approved your access to their business channell on The Nexus. This could take anything upto a working week. Without approval you cannot access the software and start selling this companies resources.</strong></p>
        <p class="style36"><strong>When you have approval you may have to sign a declaration of the company or some for of employment access agreement. As you are responsible for your own action any malicious damage will be prosecuted by law to the full extent. </strong></p>
        <p align="center" class="style35"><a href="login.php" class="style3">( Click here to return to the log-on screen ) </a></p>
      </blockquote></td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFCC33">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<?php
	}
}
}
	?>
  </table>
  
</div>
</body>
</html>
<?php

?>
