<?php
if(!session_id()){
  session_start();
}

if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view news feed.</a>";
	exit;
}

function EscapeChars($InputValue) 
	{
	
		
		return $InputValue;
	}

	  require('../connections/projectalpha.php'); 

if ($certk==-1) {
	   
	   
		if ($NewsID==0) {
			$hjSQL = sprintf("insert into newsfeed (VirtualID,aaHeading,`aaHeadingText-Size`,`aaHeadingText-Color`,aaDescription,aaLink1URL,aaLink1Desc,aaLink1Colour,`aaText-Size`,`aaText-Color`,`aaShift-Pause`,`aaShift-In-Effect`,`aaBackground-Color`,`aaBorder-Color`,escalation,ExpiryDate) VALUES('%d','$aaHeading','$aaHeadingTextSize','$aaHeadingTextColor','$aaDescription','$aaLink1URL','$aaLink1Desc','$aaLink1Colour','$aaTextSize','$aaTextColor','$aaShiftPause','$aaShiftInEffect','$aaBackgroundColor','$aaBorderColor','$escalation','$ExpiryDate')",$_SESSION['VirtualID']);
			$urlb = "publishnews.php";
		} else {
			$urlb = sprintf("publishnews.php?NewsID=%s&MD5A=%s",$NewsID,md5($ExpiryDate));
			$hjSQL = sprintf("update newsfeed Set aaHeading='%s', `aaHeadingText-Size`='$aaHeadingTextSize', `aaHeadingText-Color`='$aaHeadingTextColor', aaDescription='%s', aaLink1URL='%s', aaLink1Desc='%s', aaLink1Colour='$aaLink1Colour', `aaText-Size`='$aaTextSize', `aaText-Color`='$aaTextColor', `aaShift-Pause`='$aaShiftPause', `aaShift-In-Effect`='$aaShiftInEffect', `aaBackground-Color`='$aaBackgroundColor', `aaBorder-Color`='$aaBorderColor', escalation='$escalation',ExpiryDate='$ExpiryDate' where NewsID = %d",EscapeChars($aaHeading),EscapeChars($aaDescription),EscapeChars($aaLink1URL),EscapeChars($aaLink1Desc),$NewsID);
		}
		mysql_select_db($database_projectalpha, $projectalpha);
		$rshjSQL = mysql_query($hjSQL, $projectalpha) or die(mysql_error());
		$certk=0;
		
		?>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Newsfeed Editor</title>
	<META HTTP-EQUIV="refresh" CONTENT="1;URL=<?php echo $urlb; ?>">

	<?php
	
} else {
?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<style type="text/css">
<!--
.style32 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	color: #000000;
	font-size: x-small;
}
-->
</style>
<head>
<title>Newsfeed Editor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">




<style type="text/css">
<!--
body,td,th {
	color: #CCCCCC;
}
body {
	background-color: #FFFFCC;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 10px;
}
.style1 {
	font-size: xx-small;
	font-weight: bold;
	color: #0000FF;
}
a {
	font-family: Trebuchet MS, Tahoma, Arial;
	color: #33CC33;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #993366;
}
a:hover {
	text-decoration: none;
	color: #FF3333;
}
a:active {
	text-decoration: none;
	color: #FFCCCC;
}
.style12 {font-family: "Trebuchet MS", Tahoma, Arial; font-size: x-small; font-weight: bold; color: #000000; }
.style13 {color: #000000}
.style14 {color: #FF0033}
.style19 {
	color: #FF0033;
	font-family: "Trebuchet MS", Tahoma, Arial;
	font-size: medium;
}
.style21 {font-family: "Trebuchet MS", Tahoma, Arial; font-size: medium; }
.style22 {color: #000000; font-weight: bold;}
.style23 {font-family: "Trebuchet MS", Tahoma, Arial; font-size: medium; font-weight: bold; color: #000000; }
.style24 {font-size: medium}
.style25 {
	color: #FF6699;
	font-weight: bold;
}
.style26 {font-size: large}
.style27 {color: #993333}
.style30 {font-family: "Trebuchet MS", Tahoma, Arial; font-size: large;}
.style31 {font-size: large; color: #993333; }
.style29 {font-size: 12px}
-->
</style></head>

<body>
<?php include('top.php3'); 

    mysql_select_db($database_projectalpha, $projectalpha);
	$query_prim = sprintf("select DATE_ADD(NOW(), INTERVAL 32 DAY) as tmpExpiryDate, md5(decode(`Password`,'dr34mt1me')) as MD5A, bPrimary from sysops where RecID = '%d'",$_SESSION['SysopID']);
	$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
	$row_rsprim = mysql_fetch_assoc($rsprim);
	$totalRows_rsprim = mysql_num_rows($rsprim);
	
	if ($row_rsprim['bPrimary']==0) {
	
	?>
	
	
	<?php
		
	
	} else {
	
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_newsfeed = sprintf("select md5(ExpiryDate) as MD5A, newsfeed.* from newsfeed where VirtualID = '%d' and ExpiryDate >= Now()",$_SESSION['VirtualID']);
		$newsfeed = mysql_query($query_newsfeed, $projectalpha) or die(mysql_error());
		$totalRows_newsfeed = mysql_num_rows($newsfeed);
		
		if ($NewsID==0) {
				$aaHeading="<b></b>";
				$aaDescription="<br>";
				$aaLink1URL="";
				$aaLink1Colour="";
				$aaShiftPause="9000";
				$aaShiftInEffect="slide-up";
				$escalation="0";
				$ExpiryDate=$row_rsprim['tmpExpiryDate'];
				$aaLink1Desc="";
				}
				
		?>
		
		<table width="770"  border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#EEEEEE">
  <tr>
    <td colspan="2" align="center" valign="top"><?php
		if ($NewsID<>0) {
			include(sprintf("%s?NewsID=%d","http://www.projectalpha.com.au/newsfeed.php", $NewsID));
		}
	?></td>
    </tr>
  <tr>
    <td width="29%" align="center" valign="top">      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="39" colspan="2"><div align="center">
            <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="20">
              <param name="movie" value="NewsNews.swf">
              <param name="quality" value="high">
              <param name="bgcolor" value="#EEEEEE">
              <embed src="NewsNews.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="20" bgcolor="#EEEEEE"></embed>
            </object>
          </div></td>
        </tr>
		<?php
			while ($rsnewsfeed=mysql_fetch_assoc($newsfeed))
			 {
			 
			?>
        <tr>
          <td width="11%" height="18">		<div align="center" class="style1"><img src="images/bullet_b.gif" width="18" height="15">
		    </div>            </td>
			<?php 
			if (substr($PHP_SELF,1,5)=='https') 
			{ ?>
          		<td width="89%" bgcolor="#EEECCC"><div align="justify"><span class="style1"><a href="<?php echo sprintf("%s?NewsID=%d&MD5A=%s",$PHP_SELF,$rsnewsfeed['NewsID'],md5($rsnewsfeed['ExpiryDate'])); ?>"><?php echo $rsnewsfeed['aaHeading']; ?></a></span></div></td>
			<?php 
				} else { 
				?>
				<td width="89%" bgcolor="#EEECCC"><div align="justify"><span class="style1"><a href="<?php echo sprintf("%s?NewsID=%d&MD5A=%s",$PHP_SELF,$rsnewsfeed['NewsID'],md5($rsnewsfeed['ExpiryDate'])); ?>"><?php echo $rsnewsfeed['aaHeading']; ?></a></span></div></td>
				<?php } ?>
        </tr>
		<?php 
			
			if ($NewsID==$rsnewsfeed['NewsID'] && $MD5A==$rsnewsfeed['MD5A']) {
				$aaHeading=$rsnewsfeed['aaHeading'];
				$aaHeadingTextSize=$rsnewsfeed['aaHeadingText-Size'];
				$aaHeadingTextColor=$rsnewsfeed['aaHeadingText-Color'];
				$aaDescription=$rsnewsfeed['aaDescription'];
				$aaLink1URL=$rsnewsfeed['aaLink1URL'];
				$aaLink1Desc=$rsnewsfeed['aaLink1Desc'];
				$aaLink1Colour=$rsnewsfeed['aaLink1Colour'];
				$aaTextSize=$rsnewsfeed['aaText-Size'];
				$aaTextColor=$rsnewsfeed['aaText-Color'];
				$aaShiftPause=$rsnewsfeed['aaShift-Pause'];
				$aaShiftInEffect=$rsnewsfeed['aaShift-In-Effect'];
				$aaBackgroundColor=$rsnewsfeed['aaBackground-Color'];
				$aaBorderColor=$rsnewsfeed['aaBorder-Color'];
				$escalation=$rsnewsfeed['escalation'];
				$ExpiryDate=$rsnewsfeed['ExpiryDate'];
				$certa=0;
			}
			}
		
		}
		?>
        <tr>
          <td colspan="2"></td>
        </tr>
    </table></td>
    <td width="71%" bgcolor="#FFFFFF"><form name="form1" method="post" action=""><table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr align="left" valign="top">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr align="left" valign="top">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr align="left" valign="top" class="style23">
        <td width="30%"><span class="style12 style26 style27"><span class="style30">News ID: </span></span></td>
        <td width="70%"><span class="style31">
          <?php
			if ($NewsID==0) {
				echo 'Create New News Item, select your level';
			} else {
				echo $NewsID;
			}
		?>
        </span></td>
      </tr>
      <tr align="left" valign="top">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">News Heading</span></td>
        <td>
          <span class="style21">
          <input name="aaHeading" type="text" class="style23" id="aaHeading" size="30" maxlength="128" value="<?php echo $aaHeading; ?>">
          <span class="style14">* min 30 chars </span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Heading Text Size</span></td>
        <td><span class="style21">
          <input name="aaHeadingTextSize" type="text" class="style23" id="aaHeadingTextSize" size="3" maxlength="3" value="<?php echo $aaHeadingTextSize; ?>">
          <span class="style14">* number of points font is</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Heading Text Colour ie. ffffff </span></td>
        <td><span class="style21">
          <input name="aaHeadingTextColor" type="text" class="style23" id="aaHeadingTextColor" size="6" maxlength="6" value="<?php echo $aaHeadingTextColor; ?>">
          <span class="style14">*  absolute 6 hexadecimal chars </span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Description</span></td>
        <td><div align="center" class="style21 style24">
          <div align="left"><span class="style22">
              <textarea name="aaDescription" cols="45" rows="5" class="style23" id="aaDescription"><?php echo $aaDescription; ?></textarea>
              <br>
              * Require Content For News Tile
(min 100 chars)        </span></div>
        </div></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">End URL:</span></td>
        <td><input name="aaLink1URL" type="text" class="style23" id="aaLink1URL" size="50" value="<?php echo $aaLink1URL; ?>"></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">URLTitle: </span></td>
        <td><input name="aaLink1Desc" type="text" class="style23" id="aaLink1Desc" size="50" maxlength="128" value="<?php echo $aaLink1Desc; ?>"></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">URL Colour: </span></td>
        <td><input name="aaLink1Colour" type="text" class="style23" id="aaLink1Colour" size="6" maxlength="6" value="<?php echo $aaLink1Colour; ?>"></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Text Size: </span></td>
        <td><span class="style21">
          <input name="aaTextSize" type="text" class="style23" id="aaTextSize" size="3" maxlength="3" value="<?php echo $aaTextSize; ?>">
          <span class="style14">* number of points font is</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Text Colour: </span></td>
        <td><span class="style21">
          <input name="aaTextColor" type="text" class="style23" id="aaTextColor" size="6" maxlength="6" value="<?php echo $aaTextColor; ?>">
          <span class="style14">          *  absolute 6 hexadecimal chars</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Milliseconds to pause for: </span></td>
        <td><span class="style23">
          <input name="aaShiftPause" type="text" class="style23" id="aaShiftPause" size="10" maxlength="10" value="<?php echo $aaShiftPause; ?>">
          </span><span class="style19">* ms</span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Shift In Effect</span></td>
        <td><span class="style23">
          <select name="aaShiftInEffect" class="style23" id="aaShiftInEffect"   value="">
              <option <?php if ( $aaShiftInEffect=='slide-left') { 
			  echo 'selected'; 
			  }?> value="slide-left">slide left</option>
              <option <?php if ( $aaShiftInEffect=='slide-down') {
			  
			   echo 'selected'; 
			   }?> value="slide-down">slide down</option>
              <option <?php if ( $aaShiftInEffect=='slide-up') { 
			  echo 'selected'; 
			  }?> value="slide-up">slide up</option>
              <option <?php if ( $aaShiftInEffect=='slide-right') { 
			  echo 'selected'; 
			  }?> value="slide-right">slide right</option>
          </select>
        </span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Background Colour<br> 
          ie. cd86ef</span></td>
        <td><span class="style21">
          <input name="aaBackgroundColor" type="text" class="style23" id="aaBackgroundColor" size="6" maxlength="6" value="<?php echo $aaBackgroundColor; ?>">
          <span class="style14">*  absolute 6 hexadecimal chars</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Border Colour</span></td>
        <td><input name="aaBorderColor" type="text" class="style23" id="aaBorderColor" size="6" maxlength="6" value="<?php echo $aaBorderColor; ?>">
          <span class="style21"><span class="style14">* absolute 6 hexadecimal chars</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">Escalation</span></td>
        <td><input name="escalation" type="text" class="style23" id="escalation" size="4" maxlength="4" value="<?php echo $escalation; ?>"></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style12">ExpiryDate </span></td>
        <td><span class="style21">
          <input name="ExpiryDate" type="text" class="style23" id="ExpiryDate" size="45" maxlength="45" value="<?php echo $ExpiryDate; ?>">
          <span class="style14">*</span></span></td>
      </tr>
      <tr align="left" valign="top">
        <td><span class="style13"></span></td>
        <td><div align="center" class="style21">
          <p class="style13">
            <input name="certk" type="checkbox" id="certk" value="-1">
            <span class="style25">I certify that every thing here is just and true.</span></p>
          <p class="style13">
            <input name="createnews" type="submit" class="style22" id="createnews" value="Update/Create New News">
          </p>
        </div></td>
      </tr>
    </table>
        <p>&nbsp;</p>
    </form>
      <div align="justify">
        <blockquote class="style29">
          <div align="left" class="style32">
            <p>The Code for your website to display this series of escallations for your newsfeed is as follows: </p>
            <p>ASP: <strong>&lt;!--#include file = "http://www.projectalpha.com.au/newsfeed.php?nVirtualID=<?php echo $_SESSION['VirtualID']; ?>&amp;level=<?php echo $escalation; ?>"--&gt;</strong></p>
          </div>
        </blockquote>
      </div>
      <blockquote class="style32">
        <p align="left"> PHP: <strong>&lt;?php include(&quot;http://www.projectalpha.com.au/newsfeed.php?nVirtualID=<?php echo $_SESSION['VirtualID']; ?>&amp;level=<?php echo $escalation; ?>&quot;) ?&gt; </strong></p>
      </blockquote>
    <p class="style32">&nbsp;</p></td>
  </tr>
</table>

		
		<?
	}
?>

        <p>&nbsp;</p>
</body>
</html>
