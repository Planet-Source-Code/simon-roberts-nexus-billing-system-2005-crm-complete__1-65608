<?php require_once('../Connections/projectalpha.php'); ?>
<?php
if(!session_id()){
  session_start();
}

function ($txt) {
        
    $txt = str_replace(chr(92), sprintf("%s%s", chr(92),chr(92)),$txt) ;
    $txt = str_replace(chr(0), "\0",$txt) ;
    $txt = str_replace("'", "\'",$txt) ;
    $txt = str_replace(chr(34), sprintf("\%s", chr(34)),$txt) ;
    $txt = str_replace(chr(8), "\b",$txt);
    $txt = str_replace(chr(10), "\n",$txt) ;
    $txt = str_replace(chr(13), "\r",$txt) ;
    $txt = str_replace(chr(9), "\t",$txt) ;
    $txt = str_replace(chr(26), "\z",$txt) ;
	return $txt;
}

$mailsent = false;
require("CustomSql.inc.h4x0r.php");




	
if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">$front_pleaselogin</a>";
	exit;
} else {

	if ($_SESSION['step']==1) {
		if ($certa=="1") {
			if (($chklst1=="1") && ($chklst2=="1") && ($chklst3=="1") && ($chklst4=="1") && ($chklst5=="1") && ($chklst6=="1") && ($chklst7=="1")) {
				if ((!empty($companyname)) && (!empty($abn)) && (!empty($logourl))) {
					if (empty($realm)) {
						$realm = 'projectalpha.com.au';
					}
					mysql_select_db($database_projectalpha, $projectalpha);
					$sql = sprintf("insert into virtualisp (VirtualID, Description, ABN, ACN, LogoURL, BriefDesc, Realm, CreatedBy_SysopID) VALUES ('0', '%s', '%s', '%s', '%s', '%s', '%s', '%d')",( $companyname), ( $abn), ( $acn), ( $logourl), ( $htmlinc), ( $realm), $_SESSION['SysopID']);
					
					$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
					mysql_select_db($database_projectalpha, $projectalpha);
					$sql = sprintf("select RecID from virtualisp where Description = '%s' and ABN = '%s' and ACN = '%s' and LogoURL = '%s' and BriefDesc = '%s' and Realm = '%s'",( $companyname), ( $abn), ( $acn), ( $logourl), ( $htmlinc), ( $realm));					
					$rsNewVispID = mysql_query($sql, $projectalpha) or die(mysql_error());
					$row_NewVispID = mysql_fetch_assoc($rsNewVispID);
					
					if ($row_NewVispID['RecID']<>0) {
						$_SESSION['VirtualID'] = $row_NewVispID['RecID'];
						$_SESSION['INClause'] = sprintf("('%d')",$row_NewVispID['RecID']);
						mysql_select_db($database_projectalpha, $projectalpha);
						$sql = sprintf("insert into virtualisp_extended (RegStep, VirtualID) VALUES ('6', '%d')",$_SESSION['VirtualID']);
						$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
						mysql_select_db($database_projectalpha, $projectalpha);
						$sql = sprintf("update sysops set bConfirmed='-3', VirtualID = '%d' where RecID = '%d'",$_SESSION['VirtualID'],$_SESSION['SysopID']);
						$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
						$_SESSION['step']=6;
					}
					
					
				}
			}
		}
	}
	
	if ($_SESSION['step']==6) {
		if ($certf=="1") {
			mysql_select_db($database_projectalpha, $projectalpha);
			$sql = sprintf("update virtualisp_extended set RegStep = '2' where VirtualID = '%d'",$_SESSION['VirtualID']);
			$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
			$_SESSION['step']=2;
		}
	}
	
	if ($_SESSION['step']==2) {
		if ($certb=="1") {
			if ((!empty($fincontact)) && (!empty($finphone)) && (!empty($salescontact)) && (!empty($salesphone)) && (!empty($salesfax)) && (!empty($supcontact)) && (!empty($supphone)) && (!empty($supfax))) {		
				mysql_select_db($database_projectalpha, $projectalpha);
				$sql = sprintf("update virtualisp_extended set RegStep = '3', Finance_ContactName = '%s', Finance_PhoneNumber = '%s', Finance_FaxNumber = '%s', Finance_Email = '%s', Sales_ContactName = '%s', Sales_PhoneNumber = '%s', Sales_FaxNumber = '%s', Sales_Email = '%s', Admin_ContactName = '%s', Admin_PhoneNumber = '%s', Admin_FaxNumber = '%s', Admin_Email = '%s', Support_ContactName = '%s', Support_PhoneNumber = '%s', Support_FaxNumber = '%s', Support_Email = '%s' where VirtualID = '%d'",($fincontact), ($finphone), ($finfax), ($finemail), ($salescontact), ($salesphone), ($salesfax), ($salesemail), ($admincontact), ($adminphone), ($adminfax), ($adminemail), ($supcontact), ($supphone), ($supfax),  ($supemail), $_SESSION['VirtualID']);
				$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
				$_SESSION['step']=3;
			}
		}
	}
	
	if ($_SESSION['step']==3) {
		if ($certc=="1") {
			if ((!empty($typebus)) && (!empty($lenestablish)) && (!empty($dateowner)) && (!empty($bankbranch)) && (!empty($bankaccname)) && (!empty($bankaccstyle)) && (!empty($bankaccno)) && (!empty($audname)) && (!empty($audphone)) && (!empty($audemail)) && (!empty($creditlimit)) && (!empty($goldcreditlimit)) && (!empty($custblock)) && (!empty($extraperblock ))) {		
				if (floatval($custblock) / floatval($extraperblock) >= 30 / 6.5) {
					mysql_select_db($database_projectalpha, $projectalpha);
					$sql = "update virtualisp_extended set RegStep = '4',TradingType = '%s', History_EstablishedFor = '%s', History_DateCurOwnership = '%s', Financial_PublicBank_Designation = '%s', Financial_PublicBank_Account_Number = '%s', Financial_PublicBank_Account_AccountBSB = '%s', Financial_PublicBank_Account_SwiftCode = '%s', Financial_PublicBank_Account_Style  = '%s', Financial_Accountant_Name = '%s', Financial_Accountant_PhoneNumber = '%s', Financial_Accountant_FaxNumber = '%s', Financial_Accountant_Email = '%s',";
					$sql .= " Financial_CreditLimit = '%s', Financial_CreditLimit_CRC = md5('%s%s'), Financial_GlobalCreditLimit = '%s', Financial_GlobalCreditLimit_CRC = md5('%s%s'), Financial_CustomerBlock_Step = '%s', Financial_CustomerBlock_AdditionalBlock = '%s' where VirtualID = '%d'";
					$sql = sprintf( $sql,($typebus), ($lenestablish), (strtotime($dateowner)), ($bankbranch), ($bankaccno), ($bankaccbsb),($bankaccswift), ($bankaccname), ($bankaccstyle), ($audname), ($audphone), ($audfax), ($audemail), ($creditlimit), ('gopherprotection'), ($creditlimit), ($goldcreditlimit), ('gopherprotection'), ($goldcreditlimit), ($custblock),  ($extraperblock), $_SESSION['VirtualID']);
					print $sql;
					$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
					$_SESSION['step']=4;
				} else {
					$errormsg = "You have entered an amount below the minimum values allowed. You can order a customer block of 30 at $6.50 a block or greater.";
				}
			}
		}
	}
	
	if ($_SESSION['step']==4) {
		if ($certd=="1") {
			if ((!empty($dira_name)) && (!empty($dira_address)) && (!empty($dira_dob)) && (!empty($dira_id)) && (!empty($dira_idtype))) {		
				mysql_select_db($database_projectalpha, $projectalpha);
				$sql = sprintf("update virtualisp_extended set RegStep = '5', DirectorA_Name = '%s', DirectorA_Address = '%s', DirectorA_IDType = '%s', DirectorA_IDNumber = '%s', DirectorA_DOB = '%s', DirectorB_Name = '%s', DirectorB_Address = '%s', DirectorB_IDType = '%s', DirectorB_IDNumber = '%s', DirectorB_DOB = '%s', DirectorC_Name = '%s', DirectorC_Address = '%s', DirectorC_IDType = '%s', DirectorC_IDNumber = '%s', DirectorA_DOB = '%s' where VirtualID = '%d'",($dira_name), ($dira_address), ($dira_idtype), ($dira_id), ($dira_dob),($dirb_name), ($dirb_address), ($dirb_idtype), ($dirb_id), ($dirb_dob),($dirc_name), ($dirc_address), ($dirc_idtype), ($dirc_id), ($dirc_dob), $_SESSION['VirtualID']);
				$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
				$_SESSION['step']=5;
			}
		}
	}
	
	if ($_SESSION['step']==5) {
		if ($certe=="1") {
			if ((!empty($suppliera_name)) && (!empty($suppliera_contact)) && (!empty($suppliera_phone)) && (!empty($suppliera_email)) && (!empty($suppliera_address))) {		
				mysql_select_db($database_projectalpha, $projectalpha);
				$sql = "update virtualisp_extended set RegStep = '7', References_SupplierA_Name = '%s', References_SupplierA_ContactName = '%s', References_SupplierA_PhoneNumber = '%s', References_SupplierA_Email = '%s', References_SupplierA_Address = '%s', References_SupplierB_Name = '%s', References_SupplierB_ContactName = '%s', References_SupplierB_PhoneNumber = '%s', References_SupplierB_Email = '%s', References_SupplierB_Address = '%s', References_SupplierC_Name = '%s', References_SupplierC_ContactName = '%s', References_SupplierC_PhoneNumber = '%s', References_SupplierC_Email = '%s', References_SupplierC_Address = '%s', References_Client_Name = '%s', References_Client_ContactName = '%s', References_Client_PhoneNumber = '%s' where VirtualID = '%d'";
				$sql = sprintf($sql ,($suppliera_name), ($suppliera_contact), ($suppliera_phone), ($suppliera_email), ($suppliera_address), ($supplierb_name), ($supplierb_contact), ($supplierb_phone), ($supplierb_email), ($supplierb_address), ($supplierc_name), ($supplierc_contact), ($supplierc_phone), ($supplierc_email), ($supplierc_address),($client_name), ($client_contact), ($client_phone),$_SESSION['VirtualID']);
				$rsInsert = mysql_query($sql, $projectalpha) or die(mysql_error());
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_rsSysopDetails1 = sprintf("update sysops set SecurityLevel='90', bConfirmed = '-2', VirtualID = '%d', bWEBAccount = '0', bCreateSysop='-1', bPrimary='-1', bTemplates='-1', bRecievables='-1', bInvoice='-1', bExpenditure='-1', bHoldings='-1', bComm='-1', bRefund='-1', bAddCust='-1', bOwnership='-1', bAccSettings='-1', bVendors='-1' where RecID = '%d' ", $_SESSION['VirtualID'], $_SESSION['SysopID']);
				$rsSysopDetails1 = mysql_query($query_rsSysopDetails1, $projectalpha) or die(mysql_error());
				$_SESSION['bConfirmed']=-2;
				$_SESSION['step']=7;
			}
		}
	}
	
	if (!session_is_registered('step')) {
		$_SESSION['step'] = 1;
		}
	if ($_SESSION['step'] == 0) {
		$_SESSION['step'] = 1;
	}
	
	//mysql_select_db($database_projectalpha, $projectalpha);
	//$sql = sprintf("select RegStep from virtualisp_extended where VirtualID = '%d'",$_SESSION['VirtualID']);
	//$rsNewVispID = mysql_query($sql, $projectalpha) or die(mysql_error());
	//$row_NewVispID = mysql_fetch_assoc($rsNewVispID);
	//$_SESSION['step'] = $row_NewVispID['RegStep'];
}

?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Business Network Registration.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
<style type="text/css">
<!--
body {
	background-color: #CCCCFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 10px;
}
.style1 {color: #000000}
.style2 {font-size: 16px}
.style3 {font-family: "Trebuchet MS", Tahoma, Arial}
.style4 {font-size: 18px}
.style6 {font-size: 18px; color: #FFFFFF; }
.style11 {
	font-size: 14px;
	color: #FFFFFF;
	font-weight: bold;
}
.style13 {font-size: 12}
.style15 {color: #FF0000; font-size: 18px; }
.style16 {
	font-size: 12px;
	font-weight: bold;
}
.style18 {
	color: #FF0000;
	font-size: 12px;
	font-weight: bold;
}
.style20 {
	color: #0000FF;
	font-weight: bold;
}
body,td,th {
	font-family: Trebuchet MS, Tahoma, Arial;
}
.style21 {color: #000000; font-size: 18px; }
.style23 {font-size: 18px; font-weight: bold; }
.style24 {
	font-size: 10px;
	font-weight: bold;
}
.style25 {font-size: 24px}
.style26 {font-size: 14px}
.style28 {
	color: #FFFFFF;
	font-weight: bold;
}
.style29 {font-size: 12px}
-->
</style>
<body>

<div align="left">
  <?php
include("top.php3");
?>
  <table width="770" border="0" align="center" cellpadding="0" cellspacing="1" class="table_01">
  <?php
		if ($_SESSION['step']==1) {
		?>
    <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Create a Private Business Network Member </span></span></span></span></span></span></div></td>
    </tr>
    <tr>
      <td height="192" colspan="2" bgcolor="#FFFFFF"><span class="style1">
        <blockquote>
          <p>From here you register your business or company on the Business Network here in the Nexus. You will be able to access all public software for the network. We provided services including features such as Banking, Taxation Reports, Stock Control, Advertising, Billing Services, Retail Sales as well as many other features.</p>
          <p>It is against the law in australia to operate a business without a current non-de registered company of registered business. This mean you cannot register with this module if you do not have an ACN / RBN / ABN / BN held within australia. If you are an international company that is a company from outside of australia then you can use this application within the mean and laws of your country. </p>
          <p>This process all though is initially offered free of charge, it can incur cost of supply and runtime fee's that will be charged. You will have to agree to a Non Disclosure Agreement which acknowledges that it is a privilege to use this system and not a right. Any malicious damage will be prosecuted within the full extent of the law. Any cost you incur will be charged to your company. </p>
        </blockquote>
      </span></td>
    </tr>
	<?php
	}
	 if ($_SESSION['step']==2) {
	 ?> 
	 <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Step 1.2 - Your Company Primary Business Contacts <span class="style24">(3 to go)</span> </span></span></span></span></span></span></div></td>
     </tr>
	 <?php
	 }
	 if ($_SESSION['step']==3) {
	 ?> 
	 <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Step 1.3 - Financial Information, Bank Details, Credit Limit, Trading History <span class="style24">(2 to go)</span> </span></span></span></span></span></span></div></td>
     </tr>
	 <?php
	 }
	 if ($_SESSION['step']==4) {
	 ?> 
	 <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Step 1.4 - Business Holdings, Ownership, Partnership, Directorship <span class="style24">(1 to go)</span> </span></span></span></span></span></span></div></td>
     </tr>
	 <?php
	 }
	 if ($_SESSION['step']==5) {
	 ?> 
	 <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Step 1.5 - Trading, Supplier &amp; Client References <span class="style24">(last step)</span> </span></span></span></span></span></span></div></td>
     </tr>
	 <?php
		}
	 if ($_SESSION['step']==6) {
	 ?> 
	 <tr>
      <td colspan="2" bgcolor="#000000"><div align="left"><span class="style1"><span class="style2"><span class="style2"><span class="style3"><span class="style4"><span class="style6">Legal 1.0 - Mutual Non Disclosure Agreement</span></span></span></span></span></span></div></td>
     </tr>
	 <?php
	 }
	 if ($_SESSION['step']==1) {
		?>
    <tr>
      <td height="29" colspan="2" bgcolor="#666666"><p><span class="style11"> 1.0 - Company Listing Details </span></p></td>
    </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td width="180"><div align="right"></div></td>
      <td width="551" align="left">&nbsp;</td>
    </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="32"><div align="right" class="style13"><strong>Company Name: </strong></div></td>
      <form action="" method="post" name="form1" class="style13">
        <td align="left">      <blockquote>
          <p>
    <input name="companyname" type="text" id="companyname" size="50" maxlength="100">
    <span class="style15"></strong>*</span></p>
        </blockquote></td>
    
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="38"><div align="right" class="style13"><strong>ABN / RBN / BN: </strong></div></td>
      <td align="left"><blockquote>
        <p>
          <input name="abn" type="text" id="abn" size="50">
          <span class="style15">* </span> </p>
      </blockquote></td>
    </tr>
      <tr valign="middle" bgcolor="#CCCCCC">
        <td height="34"><div align="right"><strong>ACN:</strong></div></td>
        <td align="left"><blockquote>
          <p>
            <input name="acn" type="text" id="acn" size="50">
          </p>
        </blockquote></td>
      </tr>
        <tr valign="middle" bgcolor="#CCCCCC">
      <td height="64"><div align="right" class="style13"><strong>Realm/Domain:</strong></div></td>
      <td align="left"><blockquote>
        <p>
          <input name="realm" type="text" id="realm" size="50"> 
          <br>
          <span class="style18">(without http://www.) <br>
          ie.http://www.projectalpha.com.au would be entered as projectalpha.com.au</span></p>
      </blockquote></td>
    </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td><div align="right" class="style13"><strong>Logo URL: </strong></div></td>
      <td align="left"><blockquote>
        <p>
          <input name="logourl" type="text" id="logourl" value="http://" size="50"> 
          <span class="style15">*
          <br>
          <span class="style16">JPEG, GIF, TIF, PNG, BMP - recommended 300x300px with transparency </span></span></p>
      </blockquote></td>
    </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="227"><div align="right" class="style13">
        <p><strong>HTML to Include on Dossier:</strong></p>
        <blockquote>
          <p align="center">(List Raw HTML - after the body tag or list a plain realpath URL to a HTML Output file on the internet to include)</p>
        </blockquote>
      </div></td>
      <td align="left" valign="middle">
        <blockquote>
          <p name="textarea"><strong>
            <textarea name="htmlinc" cols="46" rows="10" id="htmlinc"></textarea>
</strong><strong>
          </strong>
                </p>
          </blockquote></td>
     	
    </tr>
    <tr valign="middle" bgcolor="#99CCFF">
      <td height="30" colspan="2" bgcolor="#666666" class="style11">1.1 - Information You will Require to complete this form - The Checklist</td>
    </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="334" colspan="2"><blockquote>
        <p align="justify">To complete registration of your business on this Business Network and to start trading on the network you will need to complete the following details. Remember you will not incur any fees until you have reached your first nominated block of customers (max 350). </p>
          <p align="justify">This will allow you to import the products on the business network share as well as of course create your own private list of vendors and traders to stem product templates from for your own independant world. Of course you get to say if your templates is shared. Your vendor information is stored away encrypted and secure from other companies prying eyes.</p>
          <p align="justify"><strong>Ok The Check List: </strong></p>
          <blockquote>
            <p align="justify"><strong>
              <input name="chklst1" type="checkbox" id="chklst1" value="1">
  </strong>Your Financial Departments Contact Person &amp; Phone Number.<br>
  <strong>
  <input name="chklst2" type="checkbox" id="chklst2" value="1">
  </strong>Your Sales Departments Contact Person, Phone &amp; Fax Number.<br>
  <strong>
  <input name="chklst3" type="checkbox" id="chklst3" value="1">
  </strong>Your Business' Registration Type ie. Sole Trader/Public/Private. etc.<br>
  <strong>
  <input name="chklst4" type="checkbox" id="chklst4" value="1">
  </strong>Your Business' Date of Registration &amp; Current Date of Ownership. <br>
  <strong>
  <input name="chklst5" type="checkbox" id="chklst5" value="1">
  </strong>Your Business' Public Trading Account detail ie. Bank, Branch, Account No. <br>
  <strong>
  <input name="chklst6" type="checkbox" id="chklst6" value="1">
  </strong>Your Business' Director / Partnerships details ie. Name, Address, ID, DOB.<br>
  <strong>
  <input name="chklst7" type="checkbox" id="chklst7" value="1">
  </strong>Your Business' Trading, Supplier, a Client References.<br> 
  </p>
            </blockquote>
      </blockquote></td>
      </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="64" colspan="2"><div align="center"><strong>
            <input name="Submit" type="submit" class="style2" value="Create my Virtual Business Identifier ">
            <br>
    Click here to move to the next step. </strong></div></td>
      </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td height="42" colspan="2">
            <div align="center">
              <input name="certa" type="checkbox" id="certa" value="1">
              <strong>I certify that the information provided here is currently registered with my country of citizenship. I also authorise you to publish this information in the business network allowing me full control of the advertising and personal news channels.</strong></div>            </td>
      </tr>
    <tr valign="middle" bgcolor="#CCCCCC">
      <td colspan="2"><div align="right"></div></td>
      </tr></form>
	<?php
	}
		if ($_SESSION['step']==2) {
		?>
    <tr bgcolor="#666666">
      <td height="30" colspan="2" class="style4"><div align="left"><span class="style3"><span class="style11"> 1.2 - Department/Office Contact and Email Information</span></span></div></td>
    </tr>
	<form action="" method="post" name="form2" class="style13">
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="118" align="right" valign="top" bordercolor="#0000FF"><div align="right" class="style21">Financial:<br>
        <span class="style16">(required)</span>      </div></td>
      <td align="center" valign="top" bordercolor="#0000FF"><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td width="120" bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <p align="left">Contact Name: </p>
          </div></td>
          <td width="337">
                <div align="left">
                  <input name="fincontact" type="text" class="style3" id="fincontact" size="50" maxlength="128">
                <span class="style15">                * </span>              </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <p align="left">Phone Number: </p>
          </div></td>
          <td>
              <div align="left">
    <input name="finphone" type="text" class="style3" id="finphone" size="50" maxlength="32">
    <span class="style15">  * </span> </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <p align="left">Fax Number: </p>
          </div></td>
          <td>
            <div align="left">
              <input name="finfax" type="text" class="style3" id="finfax" size="50" maxlength="32">
            </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <p align="left">Email:</p>
          </div></td>
          <td>
            <div align="left">
              <input name="finemail" type="text" class="style3" id="finemail" size="50" maxlength="255">
            </div></td>
        </tr>
      </table>      </td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="116" align="right" valign="top" bordercolor="#0000FF"><div align="right" class="style1"><span class="style4">Sales:<br>
        <span class="style16">(required)</span> </span></div></td>
      <td align="center" valign="top" bordercolor="#0000FF"><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td width="120" bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Contact Name: </div>
          </div></td>
          <td width="337">
                <div align="left">
                  <input name="salescontact" type="text" class="style3" id="salescontact" size="50" maxlength="128">
                <span class="style15">                * </span>              </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Phone Number: </div>
          </div></td>
          <td>
              <div align="left">
    <input name="salesphone" type="text" class="style3" id="salesphone" size="50" maxlength="32">
    <span class="style15">  * </span> </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Fax Number: </div>
          </div></td>
          <td>
            <div align="left">
              <input name="salesfax" type="text" class="style3" id="salesfax" size="50" maxlength="32">
              <span class="style15"> * </span></div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Email:</div>
          </div></td>
          <td>
            <div align="left">
              <input name="salesemail" type="text" class="style3" id="salesemail" size="50" maxlength="255">
            </div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="113" align="right" valign="top" bordercolor="#0000FF"><span class="style21">Administration:<br>
        </span></td>
      <td align="center" valign="top" bordercolor="#0000FF"><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td width="120" bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Contact Name: </div>
          </div></td>
          <td width="337">                <div align="left">
                  <input name="admincontact" type="text" class="style3" id="admincontact" size="50" maxlength="128">
                <span class="style15">                </span>              </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Phone Number: </div>
          </div></td>
          <td>              <div align="left">
    <input name="adminphone" type="text" class="style3" id="adminphone" size="50" maxlength="32">
    <span class="style15">  </span> </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Fax Number: </div>
          </div></td>
          <td>
            <div align="left">
              <input name="adminfax" type="text" class="style3" id="adminfax" size="50" maxlength="32">
            </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Email:</div>
          </div></td>
          <td>
            <div align="left">
              <input name="adminemail" type="text" class="style3" id="adminemail" size="50" maxlength="255">
            </div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="115" align="right" valign="top" bordercolor="#0000FF"><span class="style21">Support:<br>
        <span class="style16">(required)</span>      </span></td>
      <td align="center" valign="top" bordercolor="#0000FF"><table width="83%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td width="120" bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Contact Name: </div>
          </div></td>
          <td width="337">                <div align="left"><span class="style15">
                  <input name="supcontact" type="text" class="style3" id="supcontact" size="50" maxlength="128">
          </span><span class="style18"></span>              <span class="style15">*</span></div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Phone Number: </div>
          </div></td>
          <td>              <div align="left">
    <input name="supphone" type="text" class="style3" id="supphone" size="50" maxlength="32">
    <span class="style15"> * </span>     </div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Fax Number: </div>
          </div></td>
          <td>
            <div align="left">
              <input name="supfax" type="text" class="style3" id="supfax" size="50" maxlength="32">
              <span class="style15"> * </span></div></td>
        </tr>
        <tr>
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20">
            <div align="left">Email:</div>
          </div></td>
          <td>            <div align="left">
              <input name="supemail" type="text" class="style3" id="supemail" size="50" maxlength="255">
              <span class="style15">            *</span></div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="123" align="right" valign="top" bordercolor="#0000FF"><span class="style21"><br>
        </span></td>
      <td bordercolor="#0000FF"><div align="center">
        <blockquote>
          <p align="left">              <input name="certb" type="checkbox" id="certb" value="1">
              <strong>I certify that the information provided here is true and acurate.</strong></p>
          <p align="center"><strong>
            <input name="Submit2" type="submit" class="style15" value="Save Business Contact Information">
            <br>
            Click Here to goto step 1.3<br>
            (You Must Complete All the Steps 
            to be activated) <br>
          </strong></p>
        </blockquote>
      </div></td>
    </tr>
	</form>
	<?php
	}
		if ($_SESSION['step']==3) {
		?>
    <tr bgcolor="#666666">
      <td height="33" colspan="2" class="style11">1.3 - Public Trading  Bank Details / Credit Limit / Auditor Contact Information </td>
    </tr>
	
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td><span class="style15"><?php echo $errormsg ?></span></td>
    </tr><form name="form7" method="post" action="">
    <tr bgcolor="#CCCCCC" class="style16">
      <td height="104" align="right" valign="top" class="style4"><div align="right" class="style16"><strong>Business Type &amp; History: </strong></div></td>
      <td align="center" valign="top"><table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="middle">
          <td width="34%"><strong>Type of Business: </strong></td>
          <td width="66%">
            <select name="typebus" id="typebus"><form name="form7" method="post" action="">
              <option value="Sole Trader">Sole Trader</option>
              <option value="Partnership">Partnership</option>
              <option value="Public Co. (Ltd)">Public Co. (Ltd)</option>
              <option value="Private Co. (P/L)">Private Co. (P/L)</option>
              <option value="Association">Association</option>
              <option value="Other">Other</option>
            </select>
            <span class="style15">*</span>          </td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Length of establishment: </strong></td>
          <td><input name="lenestablish" type="text" id="lenestablish" size="19" maxlength="20">
            <span class="style15">*</span></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Date of current ownership: </strong></td>
          <td><input name="dateowner" type="text" id="dateowner" size="19" maxlength="20">
            <span class="style15">*</span></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC" class="style16">
      <td height="168" align="right" valign="top"><div align="right"><strong>Public Trade Account: </strong></div></td>
      <td align="center" valign="top"><table width="90%" border="0" cellspacing="0" cellpadding="0">
        <tr align="left" valign="middle">
          <td width="34%"><strong>Bank &amp; Branch:</strong></td>
          <td width="66%"><input name="bankbranch" type="text" id="bankbranch" size="44" maxlength="255">
            <span class="style15">*</span></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Account Name : </strong></td>
          <td><input name="bankaccname" type="text" id="bankaccname" size="44" maxlength="255">
            <span class="style15">*</span></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Account Style:</strong></td>
          <td><input name="bankaccstyle" type="text" id="bankaccstyle" size="44" maxlength="255">
            <span class="style15">*</span></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Account BSB: </strong></td>
          <td><input name="bankaccbsb" type="text" id="bankaccbsb" size="19" maxlength="6"></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Account Number: </strong></td>
          <td><input name="bankaccno" type="text" id="bankaccno" size="19" maxlength="20">
            <span class="style15">*</span></td>
        </tr>
        <tr align="left" valign="middle">
          <td><strong>Account Swift Code: </strong></td>
          <td><input name="bankaccswift" type="text" id="bankaccswift" size="19" maxlength="20"></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC" class="style16">
      <td height="129" align="right" valign="top"><strong>Auditor / Accountant Details: </strong></td>
      <td align="center" valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr align="left" valign="top">
          <td width="177" bordercolor="#CCCCCC" class="style3"><div align="right" class="style20 style1">
            <div align="left" class="style1">Contact Name: </div>
          </div></td>
          <td width="313">                <div align="left"><span class="style15">
                  <input name="audname" type="text" id="audname" size="44" maxlength="128">
          </span><span class="style18"></span>              <span class="style15">*</span></div></td>
        </tr>
        <tr align="left" valign="top">
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20 style1">
            <div align="left" class="style1">Phone Number: </div>
          </div></td>
          <td>              <div align="left">
    <input name="audphone" type="text" id="audphone" size="44" maxlength="32">
    <span class="style15"> * </span>     </div></td>
        </tr>
        <tr align="left" valign="top">
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20 style1">
            <div align="left" class="style1">Fax Number: </div>
          </div></td>
          <td>
            <div align="left">
              <input name="audfax" type="text" id="audfax" size="44" maxlength="32">
            </div></td>
        </tr>
        <tr align="left" valign="top">
          <td bordercolor="#CCCCCC" class="style3"><div align="right" class="style20 style1">
            <div align="left" class="style1">Email:</div>
          </div></td>
          <td>            <div align="left">
              <input name="audemail" type="text" id="audemail" size="44" maxlength="255">
              <span class="style15">            *</span></div></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC" class="style16">
      <td height="84" align="right" valign="top"><div align="right"><strong>Credit Limit &amp; Customer Block: </strong></div></td>
      <td align="center" valign="top"><div align="center">
        <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr align="left" valign="middle">
            <td><strong>Global Network Credit Limit: </strong></td>
            <td><input name="goldcreditlimit" type="text" id="goldcreditlimit" size="19" maxlength="6">
                <span class="style15"> *</span></td>
          </tr>
          <tr align="left" valign="middle">
            <td width="35%"><strong>Individual Network Affilate Credit Limit: </strong></td>
            <td width="65%"><input name="creditlimit" type="text" id="creditlimit" size="19" maxlength="6">              <span class="style15"> *</span></td>
          </tr>
          <tr align="left" valign="middle">
            <td><strong>Number of Customer per Block: </strong></td>
            <td><input name="custblock" type="text" id="custblock" value="30" size="19" maxlength="6">
              <span class="style15"> * </span></td>
          </tr>
          <tr align="left" valign="middle">
            <td><strong>Fee per addition block:</strong></td>
            <td><input name="extraperblock" type="text" id="extraperblock" value="6.50" size="19" maxlength="6">
              <span class="style15">*</span></td>
          </tr>
          <tr align="left" valign="middle">
            <td>&nbsp;</td>
            <td>              <span class="style15"> </span></td>
          </tr>
        </table>
      </div></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="124">&nbsp;</td>
      <td><div align="center">
        <p align="left">          <input name="certc" type="checkbox" id="certc" value="1">
          <strong>I certify that the information provided here is true and acurate.</strong></p>
        
          <input name="Submit3" type="submit" class="style23" value="Save Financial Information">
        <p>            <br>
          <strong>(Click here to go onto step 1.4) </strong></p>
        </div></td>
    </tr>
	</form>
	<?php
	}
		if ($_SESSION['step']==4) {
		?>
    <tr bgcolor="#666666">
      <td height="31" colspan="2"><span class="style11">1.4 - Directorship / Business owner details </span></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr><form name="form4" method="post" action="">
    <tr bgcolor="#CCCCCC">
      <td height="128" align="right" valign="top"><strong>Director / Owner A: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="36%"><strong>Name:</strong></td>
          <td width="64%">
            <input name="dira_name" type="text" id="dira_name" size="45" maxlength="130">
            <span class="style15"> *</span> </td>
        </tr>
        <tr>
          <td><strong>Address:</strong></td>
          <td><input name="dira_address" type="text" id="dira_address" size="45" maxlength="255">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td><strong>Date of Birth:</strong></td>
          <td><input name="dira_dob" type="text" id="dira_dob" size="20" maxlength="10">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td><strong>ID  Number: </strong></td>
          <td><input name="dira_id" type="text" id="dira_id" size="20" maxlength="255">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td><strong>ID Type: </strong></td>
          <td><select name="dira_idtype" id="dira_idtype">
            <option value="Driver Licence">Drivers Licence</option>
            <option value="Passport">Passport</option>
            <option value="Credit Card">Credit Card</option>
            <option value="Birth Certificate">Birth Certificate</option>
            <option value="Medical Instuition">Medical Instuition</option>
            <option value="Proof of Age">Proof of Age</option>
            <option value="Other">Other</option>
          </select>
            <span class="style15"> *</span></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="130" align="right" valign="top"><strong>Director / Owner</strong> <strong>B:</strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="36%"><strong>Name:</strong></td>
          <td width="64%">
            <input name="dirb_name" type="text" id="dirb_name" size="45" maxlength="130">
          </td>
        </tr>
        <tr>
          <td><strong>Address:</strong></td>
          <td><input name="dirb_address" type="text" id="dirb_address" size="45" maxlength="255"></td>
        </tr>
        <tr>
          <td><strong>Date of Birth:</strong></td>
          <td><input name="dirb_dob" type="text" id="dirb_dob" size="20" maxlength="10"></td>
        </tr>
        <tr>
          <td><strong>ID  Number: </strong></td>
          <td><input name="dirb_id" type="text" id="dirb_id" size="20" maxlength="255"></td>
        </tr>
        <tr>
          <td><strong>ID Type: </strong></td>
          <td><select name="dirb_idtype" id="dirb_idtype">
            <option value="Driver Licence">Drivers Licence</option>
            <option value="Passport">Passport</option>
            <option value="Credit Card">Credit Card</option>
            <option value="Birth Certificate">Birth Certificate</option>
            <option value="Medical Instuition">Medical Instuition</option>
            <option value="Proof of Age">Proof of Age</option>
            <option value="Other">Other</option>
          </select></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="129" align="right" valign="top"><strong>Director / Owner C: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="36%"><strong>Name:</strong></td>
          <td width="64%">
            <input name="dirc_name" type="text" id="dirc_name" size="45" maxlength="130">
          </td>
        </tr>
        <tr>
          <td><strong>Address:</strong></td>
          <td><input name="dirc_address" type="text" id="dirc_address" size="45" maxlength="255"></td>
        </tr>
        <tr>
          <td><strong>Date of Birth:</strong></td>
          <td><input name="dirc_dob" type="text" id="dirc_dob" size="20" maxlength="10"></td>
        </tr>
        <tr>
          <td><strong>ID  Number: </strong></td>
          <td><input name="dirc_id" type="text" id="dirc_id" size="20" maxlength="255"></td>
        </tr>
        <tr>
          <td><strong>ID Type: </strong></td>
          <td><select name="dirc_idtype" id="dirc_idtype">
            <option value="Driver Licence">Drivers Licence</option>
            <option value="Passport">Passport</option>
            <option value="Credit Card">Credit Card</option>
            <option value="Birth Certificate">Birth Certificate</option>
            <option value="Medical Instuition">Medical Instuition</option>
            <option value="Proof of Age">Proof of Age</option>
            <option value="Other">Other</option>
          </select></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="113">&nbsp;</td>
      <td><div align="center">
              <p align="left">
                <input name="certd" type="checkbox" id="certd" value="1">
                <strong>I certify that the information provided here is true and acurate.</strong></p>
            <p>              <input name="Submit32" type="submit" class="style4" value="Save Business Ownership Information">
              <br>
              <strong>(Click here to go onto step 1.5) <br>
            </strong></p>
      </div></td>
    </tr></form>
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<?php
	}
		if ($_SESSION['step']==5) {
		?>
    <tr>
      <td height="32" colspan="2" bgcolor="#666666" class="style11">1.5 - Trading &amp; Client References - <em>Last step </em></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td align="right" valign="top">&nbsp;</td>
      <td align="center" valign="top">&nbsp;</td>
    </tr><form name="form5" method="post" action="">
    <tr bgcolor="#CCCCCC">
      <td height="126" align="right" valign="top"><strong>Supplier Reference 1: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%" align="left"><strong>Supplier Name:</strong></td>
          <td width="64%">
            <input name="suppliera_name" type="text" id="suppliera_name" size="45" maxlength="255">
            <span class="style15"> *</span> </td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Name: </strong></td>
          <td><input name="suppliera_contact" type="text" id="suppliera_contact" size="45" maxlength="128">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Phone Number:</strong></td>
          <td><input name="suppliera_phone" type="text" id="suppliera_phone" size="45" maxlength="32">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Email Address: </strong></td>
          <td><input name="suppliera_email" type="text" id="suppliera_email" size="45" maxlength="255">
            <span class="style15"> *</span></td>
        </tr>
        <tr>
          <td align="left"><strong>Supplier Address:</strong></td>
          <td><input name="suppliera_address" type="text" id="suppliera_address" size="45" maxlength="255">
            <span class="style15"> *</span></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td align="right" valign="top"><strong>Supplier Reference 2: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%" align="left"><strong>Supplier Name:</strong></td>
          <td width="64%">
            <input name="supplierb_name" type="text" id="supplierb_name" size="45" maxlength="255">          </td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Name: </strong></td>
          <td><input name="supplierb_contact" type="text" id="supplierb_contact" size="45" maxlength="128"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Phone Number:</strong></td>
          <td><input name="supplierb_phone" type="text" id="supplierb_phone" size="45" maxlength="32"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Email Address: </strong></td>
          <td><input name="supplierb_email" type="text" id="supplierb_email" size="45" maxlength="255"></td>
        </tr>
        <tr>
          <td align="left"><strong>Supplier Address:</strong></td>
          <td><input name="supplierb_address" type="text" id="supplierb_address" size="45" maxlength="255"></td>
        </tr>
      </table><strong></strong></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="132" align="right" valign="top"><strong>Supplier Reference 3: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%" align="left"><strong>Supplier Name:</strong></td>
          <td width="64%">
            <input name="supplierc_name" type="text" id="supplierc_name" size="45" maxlength="255">          </td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Name: </strong></td>
          <td><input name="supplierc_contact" type="text" id="supplierc_contact" size="45" maxlength="128"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Phone Number:</strong></td>
          <td><input name="supplierc_phone" type="text" id="supplierc_phone" size="45" maxlength="32"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Email Address: </strong></td>
          <td><input name="supplierc_email" type="text" id="supplierc_email" size="45" maxlength="255"></td>
        </tr>
        <tr>
          <td align="left"><strong>Supplier Address:</strong></td>
          <td><input name="supplierc_address" type="text" id="supplierc_address" size="45" maxlength="255"></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="82" align="right" valign="top"><strong>Client Reference: </strong></td>
      <td align="center" valign="top"><table width="90%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%" align="left"><strong>Contact Name: </strong></td>
          <td width="64%"><input name="client_contact" type="text" id="client_contact" size="45" maxlength="128"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Phone Number:</strong></td>
          <td><input name="client_phone" type="text" id="client_phone" size="45" maxlength="32"></td>
        </tr>
        <tr>
          <td align="left"><strong>Contact Email Address: </strong></td>
          <td><input name="client_email" type="text" id="client_email" size="45" maxlength="255"></td>
        </tr>
      </table></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="106">&nbsp;</td>
      <td><p align="left">
        <input name="certe" type="checkbox" id="certe" value="1">
        <strong>I certify that the information provided here is true and acurate.</strong></p>
        <p align="center">
          <input name="Submit322" type="submit" class="style4" value="Finish &amp; List Me">
          <br>
          <strong>(Click here to finish the agreement) </strong></p>
        </td>
    </tr></form>
    <tr bgcolor="#CCCCCC">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<?php
	}
		if ($_SESSION['step']==6) {
			mysql_select_db($database_projectalpha, $projectalpha);
			$sql = sprintf("select Description, ABN , NOW() as DateAgg from virtualisp where RecID = '%d'",$_SESSION['VirtualID']);					
			$rsNewVispID = mysql_query($sql, $projectalpha) or die(mysql_error());
			$row_NewVispID = mysql_fetch_assoc($rsNewVispID);
		?>
	
    <tr>
      <td height="35" colspan="2" bgcolor="#666666" class="style11">Mutal Non Disclosure Agreement </td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td colspan="2"><blockquote>
        <p><br>
            <span class="style4">Agreement dated <?php echo date("l dS of F Y h:i:s A"); ?></span></p>
        <blockquote>
            <p>              <strong class="style4">Parties</strong><br>
              1 -  Exitstencil Press  Pty Ltd (ABN 87 096 867 775) of 4 Trinity Ave Millers Point, Sydney, 2000 &amp; 4 McCourt St, Wiley Park, NSW (Exitstencil Press); and<br>
              2 - <?php echo $row_NewVispID['Description']; ?> <?php echo $row_NewVispID['ABN']; ?>.<br>
              <br>
              <strong class="style4">Introduction</strong><br>
              A -  Exitstencil press and the Company are engaged in discussions and, potentially, contract negotiations concerning business operations, and, during such discussions, may disclose Confidential Information to each other (the &quot;Purpose&quot;).<br>
              B - The mutual objective of Exitstencil press and the Company is to provide appropriate protection for the Confidential Information and accordingly, Exitstencil press and the Company agree that the Confidential Information will only be used in accordance with the terms and conditions set out in this Agreement.</p>
            <p><strong class="style4">The parties agree</strong><br>
              1 - Interpretation<br>
              1.1 - Definitions: In this Agreement:<br>
              Agreement means this agreement and the Schedule attached;<br>
              Confidential Information means information (disclosed before or after this Agreement is signed) whether it is written, electronic, oral, visual or in any other form that:</p>
            <blockquote>
              <p align="justify">            (a) is by its nature confidential;<br>
                (b) is designated by the Disclosing Party as confidential; <br>
                (c) gives the Disclosing Party some competitive business advantage or the opportunity of obtaining such advantage or the disclosure of which could be detrimental to the interests of the Disclosing Party; or<br>
                (d) a Receiving Party knows or ought to know is confidential;
                  and includes:</p>
            </blockquote>
            <div align="justify">
              <blockquote>
                <blockquote>
                  <p>(i) -  information relating directly or indirectly to or about a party including past, existing or future financial information and/or business information, including details of costings, sales plans, marketing plans, financial plans, accounts and assets;<br>
                    (ii) - information created, discovered, developed or made known to the Receiving Party by the Disclosing Party during the period of, or arising out of, the Purpose;<br>
                    (iii) - trade secrets, procedures, designs, ideas, concepts, processes, inventions, research or development work undertaken or proposed to be done of a party;<br>
                    (iv) - technical knowledge, know how or information relating to the Software;<br>
                    (v)  - information relating to a party&rsquo;s customers and suppliers including, lists of customers and suppliers, price lists, and details of contracts and dealings with customers and suppliers and other parties;<br>
                    (vi)  - any legal proceedings taken or proposed to be taken by a party;<br>
                    (vii)  - any other information classifiable in equity as confidential information, 
                    whether owned by, licensed to or otherwise in the power, possession or control of the Disclosing Party.<br>
                    Disclosing Party means the party that discloses all or any part of the Confidential Information; <br>
                    Purpose means the purpose for which this Agreement has been executed and which is set out in Recital A above; 
                    Receiving Party means the party that receives all or any part of the Confidential Information;<br>
                  Software means computer software of any type or form in any stage of actual or anticipated research and development, including but not limited to programs and program modules, routines and subroutines, processes, algorithms, design concepts, design specifications (design notes, annotations, documentation, flowcharts, coding sheets, and the like), data or application systems code, source code, object code and load modules, programming patches and system designs</p>
                </blockquote>
              </blockquote>
            </div>
            <p class="style4">              <strong>1.2 - Interpretational rules:</strong></p>
            <blockquote>
              <p>              (a) - Headings are for convenience only and will be ignored in construing this Agreement;<br>
                (b) - a reference to a party means a person who is named as a party to, and is bound to observe the provisions of, this document; and<br>
                (c) - &ldquo;includes&rdquo; means includes without limitation.</p>
            </blockquote>
            <p><strong class="style4">2 - Disclosure and use of confidential information</strong><br>
              2.1 - Mutual obligation of confidentiality: In consideration of the disclosure of Confidential Information by either party, Exitstencil press and the Company agree to hold the Confidential Information in strict confidence and to use the Confidential Information solely for the Purpose.<br>
              2.2  - Restrictions on use: A Receiving Party must:</p>
            <blockquote>
              <p align="justify">            (a) - immediately inform the Disclosing Party if it suspects, or becomes aware of, any unauthorised use, copying or disclosure of the Confidential Information and take whatever reasonable action is required by the Disclosing Party to limit the damage caused by such unauthorised use, copying or disclosure<br>
                (b) - not directly or indirectly disclose, display, provide, transfer, or otherwise make available all or any part of the Confidential Information to any person or entity at any time during the period in which the Receiving Party has access to the Confidential Information, except as permitted by this Agreement or unless the Disclosing Party has given its prior written consent;<br>
                (c) - not make copies of the Confidential Information or any portion of the Confidential Information other than for the Purpose;<br>
                (d) - not reverse engineer, decompile or disassemble the Disclosing Party&rsquo;s Software or Confidential Information, or use or attempt to use Disclosing Party&rsquo;s Software in any form other than machine readable object code;<br>
                (e) - not directly or indirectly use the Confidential Information to compete with or damage a current or proposed business activity of the Disclosing Party as described in any Confidential Information;<br>
                (f) - not disclose any Confidential Information to any third party, except to those employees, advisors, agents or consultants of the Receiving party who:</p>
            </blockquote>
            <div align="justify">
              <blockquote>
                <blockquote>
                  <p>(i) - need to know such information in connection with the Purpose; and<br>
                    (ii) - understand the confidential nature of the Confidential Information and, prior to any disclosure being made, are bound to the Receiving Party by a similar duty of confidentiality and non-use. </p>
                </blockquote>
              </blockquote>
            </div>
            <p>2.3 - Excluded Information: This Agreement does not apply to any Confidential Information that a Receiving Party can document in writing:</p>
            <blockquote>
              <p>              (a) - is in the public domain through no fault of its own;<br>
                (b) - was properly known to it, without restriction, prior to disclosure by the Disclosing Party;<br>
                (c) - was properly disclosed to it, without restriction, by another person with the legal authority to do so;<br>
                (d) - is independently developed by a Receiving Party without use or reference to a Disclosing Party&rsquo;s Confidential Information;<br>
                (e) - is required to be disclosed pursuant to a judicial or legislative order or proceeding, provided the Receiving Party will not disclose any Confidential Information without first using its best efforts to inform the Disclosing Party of such legal requirement, giving the Disclosing Party a reasonable opportunity to contest such requirement and to the maximum extent possible, minimises the disclosure of the Confidential Information.            </p>
            </blockquote>
            <p>            2.4 - Period of Confidentiality: The obligations of confidentiality contained in this Agreement will continue for two years from the earlier of the first disclosure of Confidential Information or date of this Agreement.</p>
            <p><strong class="style4">3 - Return or destruction of confidential information </strong><br>
              3.1 - The Receiving Party must, immediately on demand by the Disclosing Party:</p>
            <blockquote>
              <p>            (a)return to the Disclosing Party all documents, reports, notes, memoranda, computer media and other material which record, contain, relate in any way to or are based on part or all of the Confidential Information (including all copies) which were provided to or obtained by the Receiving Party or prepared or made by, for or on behalf of the Receiving Party as a result of or in connection with the Confidential Information; <br>
                (b)delete entirely and permanently all of the Confidential Information from every computer disk or electronic storage facility of any type owned or used by Receiving Party; <br>
                (c)irrespective of anything else in this Agreement, cease to make use of the Confidential Information or any part of it; and<br>
                (d)confirm in writing promptly when it has complied with these obligations.</p>
            </blockquote>
            <p>3.2 - Urgent relief<br>
              3.3 - The parties acknowledge that damages may be inadequate compensation for breach of an obligation under this Agreement and, subject to the court&rsquo;s discretion, the Disclosing Party may restrain by an injunction or similar remedy, any conduct or threatened conduct which is or would be a breach of an obligation under this Agreement.</p>
            <p class="style4"><strong>4 - Disclaimers</strong></p>
            <blockquote>
              <p>(a) - The Disclosing Party provides all Confidential Information to the Receiving Party without any warranties of any kind.<br>
                (b) - Neither this Agreement or any disclosure of Confidential Information grants to the Receiving Party any right or licence under any trademark, copyright or other intellectual property right owned now or in the future by the Disclosing Party. </p>
            </blockquote>
            <p><span class="style23">5 - Notices</span><br>
              5.1 - Service Method: Any notice to or by a party under this Agreement must be in writing and signed by the sender and may be served by delivery in person or by post or transmission by facsimile to the address or number of the recipient specified on the first page of this Agreement or most recently notified by the recipient to the sender.<br>
              5.2 -  Receipt: Any notice shall be effective for the purposes of this Agreement upon delivery to the recipient or production to the sender of a facsimile transmittal confirmation report prior to 4.00 pm local time on a business day in the place in or to which the written notice is delivered or sent or otherwise at 9.00 am on the next business day following delivery or receipt.</p>
            <p><span class="style23">6 - General</span><br>
              6.1 -  Severability: Any provision of this Agreement which is invalid in any jurisdiction shall be ineffective in that jurisdiction to that extent, without invalidating or affecting the remaining provisions of this Agreement or the validity of that provision in any other jurisdiction.<br>
              6.2  - Assignment: A party shall not assign or otherwise transfer any right or liability under this Agreement without the prior written consent of each other party. <br>
              6.3 -  Counterparts: This Agreement may be executed in any number of counterparts, all of which taken together shall be deemed to constitute one and the same document.<br>
              6.4  - Entire agreement: This Agreement is the complete and exclusive statement of the mutual understanding of the parties and supersedes and cancels all previous written and oral agreements and communications with respect to the subject matter of this Agreement.<br>
              6.5 -  Waiver: Any waiver of any right, power, authority, discretion or remedy arising on any default under this Agreement must be in writing and signed by the party granting the waiver. A failure or delay in exercise, or partial exercise, of a right, power, authority, discretion or remedy created or arising on default under this Agreement does not result in a waiver of that right, power, authority, discretion or remedy.<br>
            6.6 -  Governing law and jurisdiction: This Agreement shall be governed by and construed under the law of the state of New South Wales and each party irrevocably submits to the exclusive jurisdiction of the courts of New South Wales.</p>
        </blockquote>
      </blockquote></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="116" colspan="2"><form name="form6" method="post" action=""><blockquote>
        <p align="left">
            <input name="certf" type="checkbox" id="certf" value="1">
            <strong>I certify that I have read this legal binding contract and as a representive of my company agree with it's terms and conditions and that I understand its constraints and the protection it offers to  my customer and company information. </strong></p>
      </blockquote>        
        
          <div align="center">
            <input name="Submit4" type="submit" class="style4" value="Agree to the MNDA">
          </div>
          <div align="center"><strong>(Click here to start the registration process) </strong>
          </div>
      </form>        </td>
    </tr>
	<?php
		mysql_free_result($rsNewVispID);
	
	}
		if ($_SESSION['step']==7) {
		
		?>
    <tr bgcolor="#000000">
      <td colspan="2" class="style6 style25">Finished - Welcome Aboard </td>
    </tr>
    <tr align="left" valign="middle" bgcolor="#00CCCC">
      <td height="887" colspan="2"><div align="left">
        <blockquote>
          <p align="justify"><span class="style26"><br>
            <br>
            Welcome to the Nexus' business networks. From here you will get access to our vendor point of sale, sysop control center and sales channel editor. These tools will enable you to run your business from any location around the world. Communicate and calculate commissions for all your staff (sysops). </span></p>
          <p align="justify"><span class="style26">Not only do you have your own private world of vendors, clients, plans and services, you can also import shared plan or service templates from other companies that have templates on offer. This will allow you to start ordering products for your customers from their 'shared' range. This goes down to every detail on your purchase order.</span></p>
          <p align="justify" class="style26">Your Virtual Identity Number is <b><?php echo $_SESSION['VirtualID']; ?> </b>this is your business number on our network. As you will notice you are all ready listed in the business network. This also comes with your own private newsfeed you can host on your website, this is added to the code using php or asp using an include statement, your statements in either asp or php are as follow, just cut an paste this into your code.</p>
          <div align="justify">
            <blockquote class="style29">
              <div align="left">ASP: <span class="style28">&lt;!--#include file = "http://www.projectalpha.com.au/newsfeed.php?nVirtualID=<?php echo $_SESSION['VirtualID']; ?>&amp;level=escalation"--&gt;</span></div>
            </blockquote>
          </div>
          <blockquote>
            <p align="left" class="style29">                PHP: <span class="style28">&lt;?php include(&quot;http://www.projectalpha.com.au/newsfeed.php?nVirtualID=<?php echo $_SESSION['VirtualID']; ?>&amp;level=escalation&quot;) ?&gt; </span></p>
            </blockquote>
          <p align="justify" class="style26">Your newsfeed will display upto the minute information on price change and details published news that you may want to run. It is a live feed from the database to the feed location of your choice.</p>
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="33%" align="center" valign="top" bordercolor="#000000"><img src="../images/photos/jarrett_sml.gif" width="195" height="269"></td>
              <td width="67%" align="left" valign="top"><p align="justify" class="style26">As one of the directors of Exitstencil Press I would like to invite you to explore our range of tools and utilities to make your business unique within the every busy day of the corperate, commerical and retail environment.</p>
                <p align="justify" class="style26">We have work hard to put together this pelimanary system for you to utilise and enjoy. We understand that no solution no matter how taylored it perfect that why we would like to have as much feed back as possible. If you think of anything that is need, something that needs change then please bring it up in our forum.</p>
                <p align="justify" class="style26">If you have any questions please do not hesitate to contact us at your convenient moment. </p>
                <p align="justify" class="style26">Kind Regards</p>
                <p class="style26">&nbsp;</p>
                <p class="style26">Jarrett Costi<br>
                  Director<br>
                  Exitstencil Press Pty Ltd </p>                </td>
            </tr>
          </table>
          <p class="style26">&nbsp; </p>
        </blockquote>
      </div></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<?php } ?>
  </table>
    

</div>
</body>
</html>
