<?php
require("./CustomSql.inc.php");
$db = new CustomSQL($DBName);
include("./usercheck.php");

if (empty($page)){
$page = 0;
}
$record = 20;

if ($Deluser==$admin_yes) {
$db->deluser($cid);
}

if (!empty($adduser)) {
$db->adduser($username,$password,$email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography);   
}

if (!empty($edituser)) {   
$db->edituser($password,$email,$homepage,$icq,$aol,$yahoo,$msn,$location,$occupation,$interests,$biography,$cid);   
}

$result = $db->getalluser($page,$record);
?>
<html>
<head>
<title><?php print "$admin_useradmin"; ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=<?php print "$admin_charset"; ?>">
<link rel="stylesheet" href="style/style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
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
              <tr bgcolor="#CCCCCC"> 
                <td>&nbsp;</td>
                <td><?php print "$admin_username"; ?></td>                               
                <td colspan="2"><?php print "$admin_opreation"; ?></td>
              </tr>
              <?php
              if (!empty($result)) {
	        while ( list($key,$val)=each($result) ) {
	        $customerid = stripslashes($val["customerid"]);
	        $username = stripslashes($val["username"]);	        
              ?>
              <tr bgcolor="#FFFFFF">
              <td><?php print "$customerid"; ?></td>
                <td><?php print "$username"; ?></td>                               
                <td><a href="edituser.php?cid=<?php print "$customerid"; ?>" class="en_b"><?php print "$admin_edit"; ?></a></td>
                <td><a href="deluser.php?cid=<?php print "$customerid"; ?>" class="en_b"><?php print "$admin_del"; ?></a></td>                              
              </tr>
              <?php
              }
              }
              ?>                       
            <tr bgcolor="#FFFFFF">
            <td align="right" colspan="4">
            <?php
              $pagenext = $page+1;
		$result1 = $db->getalluser($pagenext,$record);
		if ($page!=0)
		{
		$pagepre = $page-1;		
		print "<a href=\"$PHP_SELF?page=$pagepre\"><font color=\"#FF0000\">$admin_previouspage</font></a>&nbsp;&nbsp;&nbsp;";
		}
		if (!empty($result1))
		{
		print "<a href=\"$PHP_SELF?page=$pagenext\"><font color=\"#FF0000\">$admin_nextpage</font></a>&nbsp;";
		}
		?>
            </td>
            </tr>
            </table>            
            </td>
        </tr>    
        <tr>
        <td align="center">
        <form action="<?php print "$PHP_SELF"; ?>" method="POST" ENCTYPE="multipart/form-data">               
        <table width="300" border="0" cellspacing="1" cellpadding="4" bgcolor="#F2F2F2">
             <tr bgcolor="#FFFFFF"> 
                <td width="83"><?php print "$admin_username"; ?> :</td>
                <td width="198"><input type="text" name="username"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_password"; ?> :</td>
                <td><input type="text" name="password"></td>
              </tr>                 
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_email"; ?> :</td>
                <td><input type="text" name="email"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_homepage"; ?> :</td>
                <td><input type="text" name="homepage"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_icq"; ?> :</td>
                <td><input type="text" name="icq"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_aol"; ?> :</td>
                <td><input type="text" name="aol"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_yahoo"; ?> :</td>
                <td><input type="text" name="yahoo"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_msn"; ?> :</td>
                <td><input type="text" name="msn"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_location"; ?> :</td>
                <td><input type="text" name="location"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_occupation"; ?> :</td>
                <td><input type="text" name="occupation"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_interests"; ?> :</td>
                <td><input type="text" name="interests"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td><?php print "$admin_biography"; ?> :</td>
                <td><input type="text" name="biography"></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td>&nbsp;</td>
                <td><input type="submit" name="adduser" value="<?php print "$admin_add"; ?>"></td>
              </tr>
        </table>
        <p><a href="admin_index.php"><?php print "$admin_back"; ?></a>
            </p>
        </form>
        </td>
        </tr>    
      </table>
      
    </td>
</tr>
<tr>
    <td align="center" valign="top" height="40">&nbsp;</td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>
