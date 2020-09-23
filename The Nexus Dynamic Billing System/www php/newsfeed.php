<?php require_once('Connections/projectalpha.php'); ?>
<?php
mysql_select_db($database_projectalpha, $projectalpha);

	if (!empty($NewsID)) {
	
		$query_newsfeed =  sprintf("select * from newsfeed where NewsID = %s",$NewsID);
	} else {
		if (empty($level)) {
			$level=0;
		}
			
		$query_newsfeed =  sprintf("select * from newsfeed where ExpiryDate>=Now() and VirtualID = %d and escalation = %d",$nVirtualID,$level);
	}
//} else {
//	$query_newsfeed =  sprintf("select * from newsfeed where escalation = %d",$level);
//}
$newsfeeda = mysql_query($query_newsfeed, $projectalpha) or die(mysql_error()); 
$newsfeedb = mysql_query($query_newsfeed, $projectalpha) or die(mysql_error()); 
$totalRows_newsfeed = mysql_num_rows($newsfeedb); 

?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>The Nexus News Feed - Escalation level <?php echo sprintf("%d",$level); ?> </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body>
<table width="100%" height="100"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="99%"><APPLET code="AcuteShifter.class" codebase="http://www.projectalpha.com.au/" name="labPanelApplet"
           width="100%" height="110" align="middle" archive="AcuteShifter.jar">
          <!-- The message that will be displayed in the applet. -->
          <PARAM name="Message" value="
       <?php 
	     while ($newx = mysql_fetch_assoc($newsfeeda)) {
        ?>
 		 <NEWS<?php echo sprintf('%d',$newx['NewsID']); ?>><aah<?php echo sprintf('%d',$newx['NewsID']); ?>><?php echo sprintf('%s',$newx['aaHeading']); ?></aah<?php echo sprintf('%d',$newx['NewsID']); ?>>
		 <?php echo sprintf('%s',$newx['aaDescription']); ?>
		 <?php 
		 if (empty($newx['aaLink1URL'])) {
		 
		 } else {
		 	?>
		 <link URL='<?php echo sprintf('%s',$newx['aaLink1URL']); ?>' Text-Color='<?php echo sprintf('%s',$newx['aaLink1Colour']); ?>'><?php echo sprintf('%s',$newx['aaLink1Desc']); ?></link>
			<?php
		 } ?></NEWS<?php echo sprintf('%d', $newx['NewsID']); ?>>
		
		 <?php } ?>">
	     <PARAM name="Style" value="
         <default
            Shift-Pause='9000'
            Shift-In-Effect='slide-up'
            Padding-Top='4'
            Background-Image-Repeat='true'
            Background-Color='<?php 
		if (empty($bckcolour)) {
			echo "990000";
		} else {
			echo sprintf("%s",$bckcolour);
		} ?>'
            Border-Color='FFFFFF'
            Border-Type='full'>
	    
		<?php 
	     while ($newx = mysql_fetch_assoc($newsfeedb)) {
        ?><aah<?php echo sprintf('%d',$newx['NewsID']); ?> 
			Text-Size='<?php echo sprintf('%d',$newx['aaHeadingText-Size']); ?>'
            Text-Color='<?php echo sprintf('%d',$newx['aaHeadingText-Color']); ?>'
            Padding-Top='4'>
         </aah<?php echo sprintf('%d',$newx['NewsID']); ?>>

         <NEWS<?php echo sprintf('%d',$newx['NewsID']); ?> 
            Text-Size='<?php echo sprintf('%s',$newx['aaText-Size']); ?>'
            Text-Color='<?php echo sprintf('%s',$newx['aaText-Color']); ?>'
            Shift-Pause='<?php echo sprintf('%s',$newx['aaShift-Pause']); ?>'
            Shift-In-Effect='<?php echo sprintf('%s',$newx['aaShift-In-Effect']); ?>'
            Padding-Top='4'
            Background-Image-Repeat='true'
            Background-Color='<?php echo sprintf('%s',$newx['aaBackground-Color']); ?>'
            Border-Color='<?php echo sprintf('%s',$newx['aaBorder-Color']); ?>'
            Section-Header='true'>            
	     </NEWS<?php echo sprintf('%d',$newx['NewsID']); ?>><?php } ?>">
          <!-- The following parameters are used to format the applet
           area while images and input files are loaded. (Optional).-->
          <PARAM name="Loading-Text" value="Loading News...">
          <PARAM name="Loading-Text-Color" value="333333">
          <PARAM name="Loading-Background-Color" value="ffffff">
          <!-- When you register AcuteApplets you will get Domain-Keys 
           that removes the intro nag-screen. -->
          <PARAM name="Domain-Keys" value="13280,13213"/>
        </APPLET>
</td>
    <td width="0%" align="right"><img src="http://www.projectalpha.com.au/images/newsfeedccw.jpg" width="20" height="110"></td>
  </tr>
</table>
</body>
</html>
<?php
mysql_free_result($newsfeeda);
mysql_free_result($newsfeedb);
?>
