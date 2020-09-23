<?php
require("./CustomSql.inc.php");
$db = new CustomSQL($DBName);
include("./usercheck.php");
$userinfo = $db->getuserinfobyid($cid);
$password = $userinfo[0]["password"];
$email = $userinfo[0]["email"];
$homepage = $userinfo[0]["homepage"];
$icq = $userinfo[0]["icq"];
$aol = $userinfo[0]["aol"];
$yahoo = $userinfo[0]["yahoo"];
$msn = $userinfo[0]["msn"];
$location = $userinfo[0]["location"];
$occupation = $userinfo[0]["occupation"];
$interests = $userinfo[0]["interests"];
$biography = $userinfo[0]["biography"];
?>
<html>
<head>
<title><?php print "$admin_useradmin"; ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=<?php print "$admin_charset"; ?>">
<link rel="stylesheet" href="style/style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<form action="useradmin.php" method="POST">
<input type=.hidden. name="cid" value="<?php print "$cid"; ?>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
    <td align="center" valign="top"> 
      <?php
      include("top.php3");
      ?>
      <hr width="90%" size="1" noshade>
      <table width="90%" border="0" cellspacing="0" cellpadding="4" height="300">
        <tr> 
          <td align="center"> 
            <table width="400" border="0" cellspacing="1" cellpadding="4" bgcolor="#F2F2F2">
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_password"; ?> :</td>
                <td><input type="text" name="password" value="<?php print "$password"; ?>"></td>
              </tr>                 
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_email"; ?> :</td>
                <td><input type="text" name="email" value="<?php print "$email"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_homepage"; ?> :</td>
                <td><input type="text" name="homepage" value="<?php print "$homepage"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_icq"; ?> :</td>
                <td><input type="text" name="icq" value="<?php print "$icq"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_aol"; ?> :</td>
                <td><input type="text" name="aol" value="<?php print "$aol"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_yahoo"; ?> :</td>
                <td><input type="text" name="yahoo" value="<?php print "$yahoo"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_msn"; ?> :</td>
                <td><input type="text" name="msn" value="<?php print "$msn"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_location"; ?> :</td>
                <td><input type="text" name="location" value="<?php print "$location"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_occupation"; ?> :</td>
                <td><input type="text" name="occupation" value="<?php print "$occupation"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_interests"; ?> :</td>
                <td><input type="text" name="interests" value="<?php print "$interests"; ?>"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_biography"; ?> :</td>
                <td><input type="text" name="biography" value="<?php print "$biography"; ?>"></td>
              </tr>                                         
            </table> 
            <p>
            <input type="submit" name="edituser" value="<?php print "$admin_ok"; ?>">             
            </p>                      
          </td>
        </tr>
      </table>
      
    </td>
</tr>
<tr>
    <td align="center" valign="top" height="40">&nbsp;</td>
  </tr>
</table>
</form>
<?php
include("bottom.php3");
?>
</body>
</html>