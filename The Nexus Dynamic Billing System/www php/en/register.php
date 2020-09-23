<?php require_once('../Connections/projectalpha.php'); ?>
<?php require_once('../Connections/Epwebdev.php'); ?>

<?php

function EscapeChars($InputValue) 
	{
	
		$InputValue = str_replace( '\\', '\\\\',$InputValue); 
		
		$InputValue = str_replace(chr(0), '\\0', $InputValue); 
		
		$InputValue = str_replace("'", sprintf("%s%s",'\\',"'"), $InputValue); 
		$InputValue = str_replace( chr(34) , sprintf("%s%s",'\\' ,chr(34)),  $InputValue);
		$InputValue = str_replace( chr(8), '\\b', $InputValue);
		$InputValue = str_replace( chr(10), '\\n', $InputValue);
		$InputValue = str_replace(chr(13), '\\r', $InputValue);
		$InputValue = str_replace( chr(9), '\\t', $InputValue);
		$InputValue = str_replace( Chr(26), '\\z', $InputValue); 
		return $InputValue;
	}
	
require("CustomSql.inc.h4x0r.php");
$db = new CustomSQL($DBName);
$showtable = true;
$errortag = false;
if (!empty($adduser)) {
	
if (empty($Username)){
	$errorusername = true;
	$errortag = true;
	$errormsg = $error_Usernameempty;
}

if (empty($Password)){
	$errorpassword = true;
	$errortag = true;
	$errormsg = $error_Passwordempty;
}

if (empty($Email)){
	$erroremail = true;
	$errortag = true;
	$errormsg = $error_Emailempty;
}

if (empty($Firstname)){
	$errorfirstname = true;
	$errortag = true;
	$errormsg = $error_Passwordempty;
}

if (empty($Surname)){
	$errortag = true;
	$errorsurname = true;
	$errormsg = $error_Emailempty;
}

if (empty($Mobile)){
	$errortag = true;
	$errormobile = true;
	$errormsg = $error_Emailempty;
}

$Usernamecheckresult = $db->checkUsername($Username);
if (!empty($Usernamecheckresult)) {
	$errortag = true;
	$errormsg = $error_userexist;
}

if ($Password!=$passretype){
	$errortag = true;
	$errormsg = $error_passretypewrong;
}

if ($Email!=$Emailretype){
	$errortag = true;
	$errormsg = $error_Emailretypewrong;
}

if (!$errortag){
	$sysopid = $db->adduser(session_id(),EscapeChars($Username),EscapeChars($Password),EscapeChars($Email),EscapeChars($homepage),EscapeChars($icq),EscapeChars($aol),EscapeChars($yahoo),EscapeChars($msn),EscapeChars($location),EscapeChars($occupation),EscapeChars($interests),EscapeChars($biography),EscapeChars($Firstname),EscapeChars($Surname),EscapeChars($Description),0,0,EscapeChars($Home),EscapeChars($Work),EscapeChars($Mobile),EscapeChars($AccountNo),EscapeChars($BSB),EscapeChars($bPayNo),EscapeChars($Street1),EscapeChars($Street2),EscapeChars($Suburb),EscapeChars($Postcode),EscapeChars($State),EscapeChars($Country));
	$showtable = false;
	mysql_select_db($database_projectalpha, $projectalpha);
	$query_sysopdetails = "select RecID from sysops Where Username='$Username' and Email = '$Email' and decode(`Password`,'dr34mt1me') = '$Password'";
	$sysopdetails = mysql_query($query_sysopdetails, $projectalpha) or die(mysql_error());
	$row_sysopdetails = mysql_fetch_assoc($sysopdetails);
	$totalRows_sysopdetails = mysql_num_rows($sysopdetails);
	?>
	<HEAD>
	<META HTTP-EQUIV="refresh" CONTENT="1;URL=<?php echo sprintf("registconfirm.php?nSysopID=%d",$row_sysopdetails['RecID']); ?>">
	</HEAD>
	<?php
	mysql_free_result($sysopdetails);
	exit;
}

}

?>

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
.style1 {font-size: 18px}
.style2 {font-size: 24px}
.style3 {font-weight: bold}
.style4 {color: #FF0000}
.style5 {color: #666666}
.style6 {
	font-size: large;
	color: #FF9933;
}
.style8 {
	font-size: medium;
	color: #FF9933;
	font-family: "Trebuchet MS", Tahoma, Arial;
}
.style13 {
	font-size: small;
	font-style: italic;
	font-weight: bold;
}
.style14 {color: #FF9933}
.style17 {font-family: "Trebuchet MS", Tahoma, Arial; color: #FF9933;}
.style18 {font-size: small; color: #FF9933; font-family: "Trebuchet MS", Tahoma, Arial; font-weight: bold; }
.style19 {font-size: small; font-weight: bold; }
.style21 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	color: #FF9933;
	font-style: italic;
	font-weight: bold;
	font-size: 24px;
}
.style22 {
	font-size: 12px;
	font-weight: bold;
}
.style24 {
	font-weight: bold;
	font-size: medium;
	color: #FFFFFF;
}
.style25 {font-size: 9px}
.style26 {color: #CCCCFF}
.style29 {font-size: small; font-weight: bold; color: #669999; }
.style30 {color: #669999}
.style31 {font-family: "Trebuchet MS", Tahoma, Arial}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<?php
include("top.php3");
?>
<table width="770" border="0" cellspacing="1" cellpadding="0" align="center" class="table_01">
  <tr> 
    <td class="table_02" width="177" valign="top"> 
    <table width="160" border="0" cellspacing="0" cellpadding="4">
        <tr> 
          <td><div align="center"><img src="http://www.projectalpha.com.au/images/idents/ep_ident.jpg" width="164" height="170"></div></td>
        </tr>
        <tr> 
          <td><div align="center"></div></td>
        </tr>
      </table>
    <div align="center"></div></td>
      <td class="menu" bgcolor="#FFFFFF" valign="top" width="605"> 
     <table border="0" cellspacing="0" cellpadding="4" width="583">
        <tr> 
          <td width="575" bgcolor="#F2F2F2" class="menu_in">::<?php print "$front_registration"; ?>&nbsp;&nbsp;&nbsp;<font color="#FF0000">*</font>&nbsp;<?php print "$front_requiredinfo"; ?></td>
        </tr>
        <tr> 
          <td><?php
          if ($showtable){
             
		  	if ($errorusername){?>
            <blockquote>            </p>            <blockquote>
              <p>&nbsp;</p>
            </blockquote>
            <table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter in a valid username into the field called <span class="style5">Username</span></span></strong></td>
            </tr>
			<?php
			}
		  	if ($errorpassword){ ?>
			</table>            
		  			<table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter in a valid password into the field called <span class="style5">Password</span></span></strong></td>
            </tr>
			<?php
			}
		  	if ($erroremail){ ?>
          </table>            
			<table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter in a valid email address into the field called <span class="style5">POP3 Email Address. </span></span></strong></td>
            </tr>
						<?php
			}
		  	if ($errorfirstname){ ?>
          </table>            
			<table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter your first given or prefered name into the field called <span class="style5">Firstname</span></span></strong></td>
            </tr>
			<?php
			}
		  	if ($errorsurname){ ?>
          </table>            
			<table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter your surname into the field called <span class="style5">Surname</span></span></strong></td>
            </tr>
          </table>           
		   <?php
			}
		  	if ($errormobile){ ?>			<table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4">You need to enter in a valid mobile phone number including country code into the field called <span class="style5">Mobile Number </span></span></strong></td>
            </tr>
          </table>            
					   <?php
			} ?>
            <br>
            Please complete this registration form as best you can. You only have to complete the field with an * next to them in red. Banking details are only needed in the event of payment of commission and revenue generation with The nexus. They can also be used to make perodical payment and direct debit transactions. <br>            <form action="<?php print "$PHP_SELF"; ?>" method="POST">
              <table border=0 cellpadding=2 cellspacing=2>
<tr>
  <td width="120">Username:</td><td><input name="Username" type="text" class="txtbox" value="<?php print "$Username"; ?>" size="45">
    <span class="style1"><span class="style2">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></span></td>
</tr>
<tr>
  <td>Password: </td><td><input name="Password" type="Password" class="txtbox" value="" size="45">
    <span class="style2">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></td>
</tr>
<tr>
  <td>Comfirm Password: </td><td><input name="passretype" type="Password" class="txtbox" value="" size="45">
    <span class="style2">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></td>
</tr>
<tr>
  <td>POP3 Email address:: </td><td><input name="Email" type="text" class="txtbox" value="<?php print "$Email"; ?>" size="45">
    <span class="style2">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></td>
</tr>
<tr>
  <td>Comfirm POP3: </td><td><input name="Emailretype" type="text" class="txtbox" value="" size="45">
    <span class="style2">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></td>
</tr>
<tr>
  <td>Description:</td>
  <td><textarea name="Description" cols="45" rows="8" class="txtbox" id="Description"><?php print "$Description"; ?></textarea></td>
</tr>
<tr>
  <td>Firstname:</td>
  <td><input name="Firstname" type="text" class="txtbox" id="Firstname" value="<?php print "$Firstname"; ?>" size="45"> 
    <span class="style2">&nbsp;<font color="#FF0000">*</font></span></td>
</tr>
<tr>
  <td>Surname:</td>
  <td><input name="Surname" type="text" class="txtbox" id="Surname" value="<?php print "$Surname"; ?>" size="45">
    <span class="style2">&nbsp;<font color="#FF0000">*</font></span></td>
</tr>
<tr>
  <td>Home Number:</td>
  <td><input name="Home" type="text" class="txtbox" id="Home" value="<?php print "$Home"; ?>" size="45"></td>
</tr>
<tr>
  <td>Work Number: </td>
  <td><input name="Work" type="text" class="txtbox" id="Work" value="<?php print "$Work"; ?>" size="45"></td>
</tr>
<tr>
  <td>Mobile Number:</td>
  <td><input name="Mobile" type="text" class="txtbox" id="Mobile" value="<?php print "$Mobile"; ?>" size="45"> 
    <span class="style2">&nbsp;<font color="#FF0000">*</font></span></td>
</tr>
<tr>
  <td colspan="2"><strong><br>
    Complete for reaccuring revenue payments from software avenues: </strong></td>
  </tr>
<tr>
  <td>Bank Account Number: </td>
  <td><input name="AccountNo" type="text" class="txtbox" id="AccountNo" value="<?php print "$AccountNo"; ?>"> </td>
</tr>
<tr>
  <td>Bank BSB: </td>
  <td><input name="BSB" type="text" class="txtbox" id="BSB" value="<?php print "$BSB"; ?>"></td>
</tr>
<tr>
  <td>Bank b-Pay Number: </td>
  <td><input name="bPayNo" type="text" class="txtbox" id="bPayNo" value="<?php print "$bPayNo"; ?>"></td>
</tr>
<tr>
  <td colspan="2"><br>
    <strong>    Personal Mailling address for commission reports and statements. </strong></td>
  </tr>
<tr>
  <td>Street Line 1: </td>
  <td><input name="Street1" type="text" class="txtbox" id="Street1" value="<?php print "$Street1"; ?>" size="45"></td>
</tr>
<tr>
  <td>Street Line 2 </td>
  <td><input name="Street2" type="text" class="txtbox" id="Street2" value="<?php print "$Street2"; ?>" size="45"></td>
</tr>
<tr>
  <td>Suburb, City, Mountain:</td>
  <td><input name="Suburb" type="text" class="txtbox" id="Suburb" value="<?php print "$Suburb"; ?>" size="45"></td>
</tr>
<tr>
  <td>Postcode:</td>
  <td><input name="Postcode" type="text" class="txtbox" id="Postcode" value="<?php print "$Postcode"; ?>" size="45"></td>
</tr>
<tr>
  <td>Country:</td>
  <td><input name="Country" type="text" class="txtbox" id="Country" value="<?php print "$Country"; ?>" size="45"></td>
</tr>
<tr>
  <td colspan="2"><br>
    <strong>Personal Contact information and details: </strong></td>
  </tr>
<tr><td><?php print "$front_homepage"; ?> : </td><td><input name="homepage" type="text" class="txtbox" value="<?php print "$homepage"; ?>" size="45"></td></tr>
<tr><td><?php print "$front_icq"; ?> : </td><td><input name="icq" type="text" class="txtbox" value="<?php print "$icq"; ?>" size="45"></td></tr>
<tr><td><?php print "$front_aol"; ?> : </td><td><input name="aol" type="text" class="txtbox" value="<?php print "$aol"; ?>" size="45"></td></tr>
<tr><td><?php print "$front_yahoo"; ?> : </td><td><input name="yahoo" type="text" class="txtbox" value="<?php print "$yahoo"; ?>" size="45"></td></tr>
<tr><td><?php print "$front_location"; ?> : </td><td><textarea name="location" cols="45" rows="8" class="txtbox"><?php print "$location"; ?></textarea></td></tr>
<tr><td><?php print "$front_occupation"; ?> : </td>
  <td><textarea name="occupation" cols="45" rows="8" class="txtbox"><?php print "$occupation"; ?></textarea></td>
</tr>
<tr><td><?php print "$front_interests"; ?> : </td><td><textarea name="interests" cols="45" rows="8" class="txtbox"><?php print "$interests"; ?></textarea></td></tr>
<tr><td><?php print "$front_biography"; ?> : </td><td><textarea name="biography" cols="45" rows="8" class="txtbox"><?php print "$biography"; ?></textarea></td></tr>
<tr><td></td>
  <td>
      <div align="center">
        <input name="adduser" type="submit" class="txtbox" value="<?php print "$front_regsubmit"; ?>">
      </div></td></tr>
</table>
</form>	
          <?php
         
		}  
		  ?>        </td>
        </tr>
        <tr> 
          <td align="right">&nbsp; </td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>

