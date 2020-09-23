<?php
if(!session_id()){
  session_start();
}
require("CustomSql.inc.h4x0r.php");

if (!(session_is_registered("SysopID"))){
	?>
<HEAD>
<META HTTP-EQUIV="refresh" CONTENT="1;URL=login.php">
</HEAD>
<?php
}

$db = new CustomSQL($DBName);
$showtable = true;
$errortag = false;
if (!empty($edituser)) {
	
if (empty($email)){
	$errortag = true;
	$errormsg = $error_emailempty;
}

if ($email!=$emailretype){
	$errortag = true;
	$errormsg = $error_emailretypewrong;
}

if (!$errortag){
	$db->edituser($email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$SysopID,$Firstname,$Surname,$Description,$Checked,$VirtualID,$Home,$Work,$Mobile,$AccountNo,$BSB,$bPayNo,$Street1,$Street2,$Suburb,$Postcode,$State,$Country);
	$showtable = false;
}

}

$userinfo = $db->getuserinfobyid($SysopID);
$email = $userinfo[0]["Email"];
$homepage = $userinfo[0]["homepage"];
$icq = $userinfo[0]["icq"];
$aol = $userinfo[0]["aol"];
$yahoo = $userinfo[0]["yahoo"];
$msn = $userinfo[0]["msn"];
$location = $userinfo[0]["location"];
$occupation = $userinfo[0]["occupation"];
$interests = $userinfo[0]["interests"];
$biography = $userinfo[0]["biography"];
$Firstname = $userinfo[0]["Firstname"];
$Surname = $userinfo[0]["Surname"];
$Description = $userinfo[0]["Description"];
$Checked = $userinfo[0]["Checked"];
$VirtualID = $userinfo[0]["VirtualID"];
$Home = $userinfo[0]["Home"];
$Work = $userinfo[0]["Work"];
$Mobile = $userinfo[0]["Mobile"];
$AccountNo = $userinfo[0]["AccountNo"];
$BSB = $userinfo[0]["BSB"];
$bPayNo = $userinfo[0]["bPayNo"];
$Street1 = $userinfo[0]["Street1"];
$Street2 = $userinfo[0]["Street2"];
$Suburb = $userinfo[0]["Suburb"];
$Postcode = $userinfo[0]["Postcode"];
$State = $userinfo[0]["State"];
$Country = $userinfo[0]["Country"];

?>
<html>
<head>
<title><?php print "$front_modiinfo"; ?></title>
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
	background-color: #DEDECA;
}
-->
</style></head>

<body text="#000000" leftmargin="0" topmargin="0">
<?php
include("top.php3");
?>
<table width="770" border="0" cellspacing="1" cellpadding="0" align="center" class="table_01">
  <tr> 
    <td class="table_02" width="160" valign="top"> 
    <table width="160" border="0" cellspacing="0" cellpadding="4">
        <tr> 
          <td><div align="center"><img src="http://www.projectalpha.com.au/images/idents/ep_ident.jpg" width="164" height="170"></div></td>
        </tr>
        <tr> 
          <td><div align="center">ACN 096 867 775 </div></td>
        </tr>
      </table>
    </td>
      <td class="menu" bgcolor="#FFFFFF" valign="top" width="610"> 
     <table border="0" cellspacing="0" cellpadding="4" width="610">
        <tr> 
          <td bgcolor="#F2F2F2" class="menu_in">::<?php print "$front_modiinfo"; ?>&nbsp;&nbsp;&nbsp;<font color="#FF0000">*</font>&nbsp;<?php print "$front_requiredinfo"; ?></td>
        </tr>
        <?php
        if ($errortag){
        ?>
        <tr> 
          <td><font color="#FF0000"><?php print "$errormsg"; ?></font></td>
        </tr>
        <?php
        }
        ?>
        <tr> 
          <td> 
          <?php
          if ($showtable){
          ?>                     
<form action="<?php print "$PHP_SELF"; ?>" method="POST">
<table border=0 align="center" cellpadding=2 cellspacing=2>
<tr>
  <td width="120">POP3 Email address: </td><td><input type="text" name="email" value="<?php print "$email"; ?>">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr>
  <td>Confirm POP3: </td><td><input type="text" name="emailretype" value="">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr>
  <td>Firstname:</td>
  <td><input type="text" name="Firstname" value="<?php print "$Firstname"; ?>"></td>
</tr>
<tr>
  <td>Surname:</td>
  <td><input type="text" name="Surname" value="<?php print "$Surname"; ?>"></td>
</tr>
<tr>
  <td>Home Number:</td>
  <td><input name="Home" type="text" id="Home" value="<?php print "$Home"; ?>"></td>
</tr>
<tr>
  <td>Work Number: </td>
  <td><input name="Work" type="text" id="Work" value="<?php print "$Work"; ?>"></td>
</tr>
<tr>
  <td>Mobiel Number:</td>
  <td><input name="Mobile" type="text" id="Mobile" value="<?php print "$Mobile"; ?>"></td>
</tr>
<tr>
  <td colspan="2"><strong><br>
    Complete for reaccuring revenue payments from software avenues: </strong></td>
</tr>
<tr>
  <td>Bank Account Number: </td>
  <td><input name="AccountNo" type="text" id="AccountNo" value="<?php print "$AccountNo"; ?>">
  </td>
</tr>
<tr>
  <td>Bank BSB: </td>
  <td><input name="BSB" type="text" id="BSB" value="<?php print "$BSB"; ?>"></td>
</tr>
<tr>
  <td>Bank b-Pay Number: </td>
  <td><input name="bPayNo" type="text" id="bPayNo" value="<?php print "$bPayNo"; ?>"></td>
</tr>
<tr>
  <td colspan="2"><br>
      <strong> Personal Mailling address for commission reports and statements. </strong></td>
</tr>
<tr>
  <td>Street Line 1: </td>
  <td><input name="Street1" type="text" id="Street1" value="<?php print "$Street1"; ?>"></td>
</tr>
<tr>
  <td>Street Line 2 </td>
  <td><input name="Street2" type="text" id="Street2" value="<?php print "$Street2"; ?>"></td>
</tr>
<tr>
  <td>Suburb, City, Mountain:</td>
  <td><input name="Suburb" type="text" id="Suburb" value="<?php print "$Suburb"; ?>"></td>
</tr>
<tr>
  <td>Postcode:</td>
  <td><input name="Postcode" type="text" id="Postcode" value="<?php print "$Postcode"; ?>"></td>
</tr>
<tr>
  <td>Country:</td>
  <td><input name="Country" type="text" id="Country" value="<?php print "$Country"; ?>"></td>
</tr>
<tr><td><?php print "$front_homepage"; ?> : </td><td><input type="text" name="homepage" value="<?php print "$homepage"; ?>"></td></tr>
<tr>
  <td>MSN : </td>
  <td><input name="msn" type="text" id="msn" value="<?php print "$msn"; ?>"></td>
</tr>
<tr><td><?php print "$front_icq"; ?> : </td><td><input type="text" name="icq" value="<?php print "$icq"; ?>"></td></tr>
<tr><td><?php print "$front_aol"; ?> : </td><td><input type="text" name="aol" value="<?php print "$aol"; ?>"></td></tr>
<tr><td><?php print "$front_yahoo"; ?> : </td><td><input type="text" name="yahoo" value="<?php print "$yahoo"; ?>"></td></tr>
<tr><td><?php print "$front_location"; ?> : </td><td><textarea name="location"><?php print "$location"; ?></textarea></td></tr>
<tr><td><?php print "$front_occupation"; ?> : </td><td><textarea name="occupation"><?php print "$occupation"; ?></textarea></td></tr>
<tr><td><?php print "$front_interests"; ?> : </td><td><textarea name="interests"><?php print "$interests"; ?></textarea></td></tr>
<tr><td><?php print "$front_biography"; ?> : </td><td><textarea name="biography"><?php print "$biography"; ?></textarea></td></tr>
<tr><td></td><td><input type="submit" name="edituser" value="<?php print "$front_modiinfo"; ?>"></td></tr>
</table>
</form>	
          <?php
	}
	else{
	?>
	<a href="login.php"><?php print "$front_back"; ?></a>
	<?php
	}
	?>
          </td>
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
