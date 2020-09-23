<?php
if(!session_id()){
  session_start();
}

if (!(session_is_registered("SysopID"))){
	print "<a href=\"login.php\">Please login to view news feed.</a>";
	exit;
}
 require_once('../Connections/projectalpha.php'); 

?>
<?php

    mysql_select_db($database_projectalpha, $projectalpha);
	$query_prim = sprintf("select md5(decode(`Password`,'dr34mt1me')) as MD5A, bPrimary from sysops where RecID = '%d'",$_SESSION['SysopID']);
	$rsprim = mysql_query($query_prim, $projectalpha) or die(mysql_error());
	$row_rsprim = mysql_fetch_assoc($rsprim);
	$totalRows_rsprim = mysql_num_rows($rsprim);
	

	if ($row_rsprim['bPrimary']<>0) {
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsStationary = sprintf("Select stationary.RecID, stationary.StationaryCode, stationary.HTML, concat(sysops.Firstname,' ',sysops.Surname) as Fullname FROM stationary inner join sysops on stationary.LastSysopID = sysops.RecID where stationary.VirtualID = '%d'",$_SESSION['VirtualID']);
		$rsStationary = mysql_query($query_rsStationary, $projectalpha) or die(mysql_error());
		$totalRows_rsStationary = mysql_num_rows($rsStationary);
} else {
?>
Compromising the Host will not be tollerated
<?php
//exit;
}

?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<?php 
	if ($totalRows_rsStationary==0) {

		if (!empty($cmdcreate) && !empty($ccode))  {
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_update = sprintf("insert into stationary (HTML, StationaryCode, LastSysopID, SysopID, VirtualID) VALUES ('%s','%s', '%d','%d','%d')",$shtml, $ccode, $_SESSION['SysopID'],  $_SESSION['SysopID'],  $_SESSION['VirtualID']);  
				$rsupdate = mysql_query($query_update, $projectalpha) or die(mysql_error());
			}
?>
<title>The Nexus Stationary Editor: You have no stationary, how will you bill? Add one now..</title>
<?php 
} else {
?>
<title>The Nexus Stationary Editor: <?php echo $totalRows_rsStationary; ?> Stationary Templates
<?php

 while ($row_rsStationary = mysql_fetch_assoc($rsStationary))
	{

		$xhtml .= sprintf("<tr><td width=\"%s\"><img src=\"../images/icons/bullet.GIF\" width=\"22\" height=\"22\"></td><td colspan=\"2\" bordercolor=\"#000000\" bgcolor=\"#FEEFCB\"> - <a href=\"publishstationary.php?code=%s\">%s</a></td></tr>",'4%',$row_rsStationary['StationaryCode'],$row_rsStationary['StationaryCode']);
		if ($row_rsStationary['StationaryCode'] == $code && empty($ccode)) {

			if (!empty($cmdcreate) && $row_rsStationary['HTML'] <> $hhtml)  {
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_update = sprintf("update stationary set HTML = '%s', LastSysopID='%d' where RecID = '%d',",$hhtml,$_SESSION['SysopID'],$row_rsStationary['RecID']);
				$rsupdate = mysql_query($query_update, $projectalpha) or die(mysql_error());
				$code = $row_rsStationary['StationaryCode'];
				$shtm = $hhtml;
			} else {
			
				$code = $row_rsStationary['StationaryCode'];
				$shtm = $row_rsStationary['HTML'];
			}
			
		} else {
			
			if (!empty($cmdcreate) && !empty($ccode))  {
				if ($noinsert==false) {
					mysql_select_db($database_projectalpha, $projectalpha);
					$query_update = sprintf("insert into stationary (HTML, StationaryCode, LastSysopID, SysopID, VirtualID) VALUES ('%s','%s','%d','%d','%d')",$hhtml, $ccode, $_SESSION['SysopID'],  $_SESSION['SysopID'],  $_SESSION['VirtualID']);  
					$rsupdate = mysql_query($query_update, $projectalpha) or die(mysql_error());
					$noinsert=true;
				}
			}
		}
	}
} 
?>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div align="center">
  <?php include('top.php3'); ?><form name="form1" method="post" action="">
  <table width="770" border="0" cellspacing="0" cellpadding="0">
    <tr align="center" bgcolor="#FEEDCB">
      <td valign="top">&nbsp;</td>
      <td valign="top"><em><strong>HTML STATIONARY CODE</strong></em></td>
    </tr>
    <tr align="center" bgcolor="#FEEDCB">
      <td width="284" valign="top"><div align="center">
        <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <?php
				if ($totalRows_rsStationary<6) { ?>
 <td colspan="2">
                <div align="left">                 <select name="ccode" id="ccode">
					<?php if (strpos($xhtml,'INVOICE')===false) { ?>
						<option value="INVOICE">INVOICE</option>
						<?php } ?>
						<?php if (strpos($xhtml,'RECEIPT')===false) { ?>
						<option value="RECEIPT">RECEIPT</option>
						<?php } ?>
						<?php if (strpos($xhtml,'STATEMENT')===false) { ?>
						<option value="STATEMENT">STATEMENT</option>
						<?php } ?>
						<?php if (strpos($xhtml,'QUOTA')===false) { ?>
						<option value="QUOTA">QUOTA</option>
						<?php } ?>
						<?php if (strpos($xhtml,'PURCHASEORDER')===false) { ?>
						<option value="PURCHASEORDER">PURCHASEORDER</option>
						<?php } ?>
						<?php if (strpos($xhtml,'WELCOMENOTE_ONLINE')===false) { ?>
						<option value="WELCOMENOTE_ONLINE">WELCOMENOTE_ONLINE</option>
						<?php } ?>
                  </select>
				<?php } else { ?> <td >
                <div align="left">&nbsp;<?php } ?>
                </div></td>
            <td width="43%"><div align="right">
              <input name="cmdcreate" type="submit" id="cmdcreate" value="Create/Update">
            </div></td>
          </tr>
          <tr>
            <td colspan="3"></td>
          </tr>
         <?php echo $xhtml; ?>
        </table>
      </div></td>
      <td width="486" valign="top">
        <textarea name="hhtml" cols="40" rows="15" wrap="VIRTUAL" id="hhtml"><?php echo $shmt; ?></textarea></td>
    </tr>
    <tr align="center" bgcolor="#FEEDCB">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><strong>Functions List - Extra Field Taging </strong></td>
    </tr>
    <tr align="center" valign="top">
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
  </table></form>
  <br>
  <?php include('bottom.php3'); ?>
</div>
</body>
</html>
<?php
mysql_free_result($rsStationary);
?>
