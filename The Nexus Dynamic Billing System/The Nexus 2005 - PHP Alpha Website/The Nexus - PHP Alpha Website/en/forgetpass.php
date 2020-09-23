<?php
require("CustomSql.inc.h4x0r.php");
$db = new CustomSQL($DBName);
$showtable = true;
$errortag = false;

if (empty($email)){
	$errortag = true;
	$errormsg = $error_emailempty;
}

$password = $db->emailcheck($email);
if (!is_string($password)) {
	$errornoemail = true;
	$errortag = true;
	$errormsg = $error_notamember;
}

if (!$errortag){
	$showtable = false;
	$message = "Hi, \n";
	$message .= "Welcome to the Nexus, you have recieved this message because you or someone who knows your email address has entered your registered address into the The Nexus Sysop system. If this was not you there is some prevantave messure you can take. A simple one is to change your email address in the Nexus Sysop Administrator or log in online here at 'http://www.projectalpha.com.au/en/login.php'. Then click on modify info and change your email address. ";
	$message .= "
	";
	$message .= "Your password is \n";
	$message .= $password;
	$message .= "
	";
	$message .= "email: admin@projectalpha.com.au with any inquiry's you may have.
	";
	mail("$email","$front_password","$message");
}



?>
<html>
<head>
<title><?php print "$front_forgetpass"; ?></title>
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
<style type="text/css">
<!--
.style1 {font-weight: bold}
.style2 {
	color: #FF3333;
	font-weight: bold;
}
.style3 {color: #FF3333}
.style4 {font-style: italic}
-->
</style></head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
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
          <td bgcolor="#F2F2F2" class="menu_in">::<?php print "$front_forgetpass"; ?>&nbsp;&nbsp;&nbsp;<font color="#FF0000">*</font>&nbsp;<?php print "$front_requiredinfo"; ?></td>
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
            <p>Please enter your registered email address for this server. Once you have completed the email address click on the button and it will look up your password.<br>
              <?php
				if ($errornoemail) {
					?>
            </p>
            <table width="85%"  border="0">
            <tr>
              <td width="5%"><img src="images/warning.gif" width="18" height="16"></td>
              <td width="95%"><strong><span class="style4"><span class="style3">The email address enter is not on our database.</span> </span></strong></td>
            </tr>
          </table>
		  <?php		  
		  }
            if ($showtable){
			?>
		  <br>
		  <form action="<?php print "$PHP_SELF"; ?>" method="POST">
<table border=0 cellpadding=2 cellspacing=2>
<tr><td width="120"><?php print "$front_email"; ?> : </td><td><input type="text" name="email" value="<?php print "$email"; ?>">&nbsp;<font color="#FF0000">*</font></td></tr>
<tr><td><a href="register.php"><?php print "$front_registration"; ?></a>&nbsp;&nbsp;&nbsp;<a href="forgetpass.php"><?php print "$front_forgetpass"; ?></a></td><td><input type="submit" name="sendpass" value="<?php print "$front_send"; ?>"></td></tr>
</table>
</form>
	<?php
	}
	else{
	?>
			<br>
            <em><strong>The email address you have enter has been sent a message containing your password</strong></em>.<br>
            <br> 
            Please do not share your password with anyone. You can create a sysop username for them with your Nexus Sysop Administrator, if you cannot contact your admin and let them complete a sysop registration form.<br>
	<?php 
	}
	?>          </td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
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
