<?php
if(!session_id()){
  session_start();
}
require("CustomSql.inc.h4x0r.php");
if (empty($_SESSION['SysopID'])){
	print "<a href=\"login.php\">$front_pleaselogin</a>";
	exit;
}
$db = new CustomSQL($DBName);
$showtable = true;
$errortag = false;
if (!empty($modipass)) {
	
if (empty($password)){
	$errortag = true;
	$errormsg = $error_passwordempty;
}

if (empty($oldpass)){
	$errortag = true;
	$errormsg = $error_passwordempty;
}

if ($password!=$passretype){
	$errortag = true;
	$errormsg = $error_passretypewrong;
}

$checkresult = $db->checkpassword($SysopID,$oldpass);
if ($checkresult==0) {
	$errortag = true;
	$errormsg = $error_wrongpassword;
}

if (!$errortag){
	$db->modifypass($password,$SysopID);
	$showtable = false;
}

}

?>
<html>
<head>
<title><?php print "$front_modipass"; ?></title>
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
          <td><div align="center">ACN 069 867 775 </div></td>
        </tr>
      </table>
    </td>
      <td class="menu" bgcolor="#FFFFFF" valign="top" width="610"> 
     <table border="0" cellspacing="0" cellpadding="4" width="610">
        <tr> 
          <td bgcolor="#F2F2F2" class="menu_in">::<?php print "$front_modipass"; ?>&nbsp;&nbsp;&nbsp;<font color="#FF0000">*</font>&nbsp;<?php print "$front_requiredinfo"; ?></td>
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
<table border=0 cellpadding=2 cellspacing=2>
<tr><td width="120"><?php print "$front_oldpassword"; ?> : </td><td><input type="password" name="oldpass" value="">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr><td><?php print "$front_newpassword"; ?> : </td><td><input type="password" name="password" value="">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr><td><?php print "$front_passwordagain"; ?> : </td><td><input type="password" name="passretype" value="">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr><td></td><td><input type="submit" name="modipass" value="<?php print "$front_modipass"; ?>"></td></tr>
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
