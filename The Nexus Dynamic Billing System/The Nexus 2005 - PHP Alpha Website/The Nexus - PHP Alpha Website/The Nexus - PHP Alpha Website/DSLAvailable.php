<?php require('Connections/Epwebdev.php'); ?>
<?php
function GetSQLValueString($theValue, $theType, $theDefinedValue = "", $theNotDefinedValue = "") 
{
  $theValue = (!get_magic_quotes_gpc()) ? addslashes($theValue) : $theValue;

  switch ($theType) {
    case "text":
      $theValue = ($theValue != "") ? "'" . $theValue . "'" : "NULL";
      break;    
    case "long":
    case "int":
      $theValue = ($theValue != "") ? intval($theValue) : "NULL";
      break;
    case "double":
      $theValue = ($theValue != "") ? "'" . doubleval($theValue) . "'" : "NULL";
      break;
    case "date":
      $theValue = ($theValue != "") ? "'" . $theValue . "'" : "NULL";
      break;
    case "defined":
      $theValue = ($theValue != "") ? $theDefinedValue : $theNotDefinedValue;
      break;
  }
  return $theValue;
}

$editFormAction = $_SERVER['PHP_SELF'];
if (isset($_SERVER['QUERY_STRING'])) {
  $editFormAction .= "?" . htmlentities($_SERVER['QUERY_STRING']);
}

if (empty($areacode)){
	$showform = true;
	}
else
{
	$showform = false;
	$number = str_replace(' ', '' ,$phoneno);
	$areac = str_replace(' ', '' ,$areacode);
	$url = "http://www.comcen.com.au/cgi-bin/checkadslnumber3.cgi?task=find&areacode=$areacode&number=$number&extra=xml&user=ep.net.au";
	$xHandle = fopen($url,"r") ;
	$xData = fread($xHandle, 64000) ;
	fclose($xHandle);
	$xData = ereg_replace("[\r,\n]", "", $xData);
  	$number = ereg_replace(".*<number>","",$xData);
	$number = ereg_replace("</number>.*","",$number);
	$status = ereg_replace(".*<status>","",$xData);
	$status = ereg_replace("</status>.*","",$status);
	$exchange = ereg_replace(".*<exchange>","",$xData);
	$exchange = ereg_replace("</exchange>.*","",$exchange);
	$state = ereg_replace(".*<state>","",$xData);
	$state = ereg_replace("</state>.*","",$state);
	}


if ((isset($_POST["MM_insert"])) && ($_POST["MM_insert"] == "form1")) {
  $insertSQL = sprintf("INSERT INTO dslcheck (areacode, phone, username, email, accname, state, status, phonenumber, exchange) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
                       GetSQLValueString($_POST['areacode'], "text"),
                       GetSQLValueString($_POST['phoneno'], "text"),
                       GetSQLValueString($_POST['username'], "text"),
                       GetSQLValueString($_POST['email'], "text"),
                       GetSQLValueString($_POST['accname'], "text"), GetSQLValueString($state, "text"), GetSQLValueString($status, "text"), GetSQLValueString($number, "text"), GetSQLValueString($exchange, "text"));

  mysql_select_db($database_Epwebdev, $Epwebdev);
  $Result1 = mysql_query($insertSQL, $Epwebdev) or die(mysql_error());

	 $nRecID_EnquiryID = mysql_insert_id();

		mysql_select_db($database_Epwebdev, $Epwebdev);
		
		$query_EnquiryID = sprintf("SELECT dslcheck.RecID, dslcheck.`state`, dslcheck.queried FROM dslcheck WHERE dslcheck.areacode = '%s' and  dslcheck.phone = '%s'", GetSQLValueString($_POST['areacode'], "text"),GetSQLValueString($_POST['phoneno'], "text"));
		$EnquiryID = mysql_query($query_EnquiryID, $Epwebdev) or die(mysql_error());
		$row_EnquiryID = mysql_fetch_assoc($EnquiryID);
		$totalRows_EnquiryID = mysql_num_rows($EnquiryID);  

}




?>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>DSL Availability</title>
<?php
// nexpv01 - Keywords Intial Tag Inclusion
// nexpv02 - Description Intial Tag Inclusion

$nexpv01 = 'Broadband Exchange Test (Australia)';
$nexpv02 = 'Broadband Exchange Test (Australia)';
include(sprintf("http://www.projectalpha.com.au/meta.php?nexpv01='%s'&nexpv02='%s%s'",str_replace(' ','%20',$nexpv01), str_replace(' ','%20',$nexpv02))); 
?><style type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 10px;
	background-color: #FFFFCC;
}
-->
</style>
<link href="/css/txtbox.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style3 {font-family: Geneva, Arial, Helvetica, sans-serif; font-weight: bold; }
.style4 {
	font-family: Geneva, Arial, Helvetica, sans-serif;
	color: #990033;
}
.style6 {font-family: Geneva, Arial, Helvetica, sans-serif; color: #FFCCCC; font-weight: bold; }
.style8 {font-family: Geneva, Arial, Helvetica, sans-serif; color: #000000; font-weight: bold; }
.style9 {font-family: Geneva, Arial, Helvetica, sans-serif}
.style12 {
	color: #FF0000;
	font-size: 16px;
	font-weight: bold;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
.style13 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #FF0000;
	font-weight: bold;
}
.style14 {color: #FF0000}
.style15 {color: #993333}
.style16 {
	color: #990033;
	font-weight: bold;
}
.style17 {color: #990000}
-->
</style>
</head>

<body>
<table  bgcolor="#F1DA5C" width="100%"  border="0">
  <tr>
    <td><div align="right"><img src="images/free-broadband-test.jpg" width="450" height="124"></div></td>
  </tr>
</table>
<table width="80%" bgcolour="#3363FC" border="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="62%"  border="0" align="center">
  <tr>
    <td colspan="4"><div align="justify">
      <p><strong><span class="style4">Enter you local phone number in australia and test for broadband services. This will tell you if broadband is available either by turn-key satellite technologies or landland services. This is a free broadband services test for australia.</span></strong></p>
      <p class="style16">This information is kept confidential and not released to the public. It will tell you the local exchange and whether you have DSL Services Available yet. </p>
      <blockquote>
        <p><span class="style6"><span class="style15">This is available from our Reseller's who can provide this to you if your Exchange supports DSL services. see <a href="visp.net.php">Reseller Network.</a> We are currently operate on an ATM Network within australia. It comes in muliple of mode T1, DMT &amp; others.<br>
            </span><br> 
          </span></p>
      </blockquote>
    </div></td>
  </tr>
  <?php if ( $showform ) { ?> 
  <tr>
    <td colspan="4"><form name="form1" method="POST" action="<?php echo $editFormAction; ?>">
      <table width="95%"  border="0" align="center">
        <tr>
          <td><span class="style3">Sysop Username: </span></td>
          <td>&nbsp;</td>
          <td><input name="username" type="password" class="txtbox" id="username" maxlength="64"></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="style3">Valid Email Address: </span></td>
          <td>&nbsp;</td>
          <td><input name="email" type="text" class="txtbox" id="email2"></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="style3">Account Name: </span></td>
          <td>&nbsp;</td>
          <td><input name="accname" type="text" class="txtbox" id="accname2"></td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td width="36%"><span class="style3">
            <label>Area Code:</label>
&nbsp;</span></td>
          <td width="64%">&nbsp;</td>
          <td width="64%"><input name="areacode" type="text" class="txtbox" id="areacode2" size="2" maxlength="2"></td>
          <td width="64%"><div align="center" class="style12">*</div></td>
        </tr>
        <tr>
          <td><span class="style3">Landline: </span></td>
          <td>&nbsp;</td>
          <td><input name="phoneno" type="text" class="txtbox" id="phoneno2" size="8" maxlength="8"></td>
          <td><div align="center" class="style13">*</div></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><div align="right">
            <input type="submit" name="Submit" value="Search for ADSL &amp; g.SHDSL Capabilities">
          </div></td>
          <td>
            <div align="right">            </div></td>
        </tr>
      </table>
      <div align="right"><span class="style14">* Only required fields to see if your exchange supports DSL Services </span>
      </div>
    </form>
</td>

  </tr>
  	  <?php
	  } else {
		
		

	 ?>
      <tr>
        <td><div align="right" class="style9"><strong>Enquiry ID:</strong></div></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><span class="style8"><span class="style3"><?php print "$state" ?></span><span class="style3"><?php print "$number" ?></span></span></td>
      </tr>
  <tr>
    <td><div align="right" class="style9"><strong>Status:</strong></div></td>
    <td>&nbsp;</td>
    <td width="3%">&nbsp;</td>
    <td width="73%"><span class="style3">Landline <?php print "$status" ?><br>
    Satellite Broadband is also Available <span class="style17">100%</span> Coverage </span></td>
  </tr>
  <tr>
    <td><div align="right" class="style9"><strong>Number:</strong></div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><span class="style3"><?php print "$number" ?></span></td>
  </tr>
  <tr>
    <td><div align="right" class="style9"><strong>Exchange:</strong></div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><span class="style3"><?php print "$exchange" ?></span></td>
  </tr>
  <tr>
    <td width="21%"><div align="right" class="style9"><strong>State:</strong></div></td>
    <td width="3%">&nbsp;</td>
    <td>&nbsp;</td>
    <td><p class="style3"><?php print "$state" ?><br>
      </p>    </td>
  </tr>
  <?php

?>

  <?php
  }
  ?>
</table>
</body>
</html>
