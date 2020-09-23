<?php   if(strpos(" " . strtoupper($HTTP_ACCEPT),"vnd.wap.wml") > 0) {        // Check whether the browser/gateway says it accepts WML.
    //////header("Content-type: text/vnd.wap.wml");
?><head>
<title>Free Ringtones? Want a Profile?</title>
<?php
  }
  else {
    $browser=substr(trim($HTTP_USER_AGENT),0,4);
    if($browser=="Noki" ||			// Nokia phones and emulators
      $browser=="Eric" ||			// Ericsson WAP phones and emulators
      $browser=="WapI" ||			// Ericsson WapIDE 2.0
      $browser=="MC21" ||			// Ericsson MC218
      $browser=="AUR " ||			// Ericsson R320
      $browser=="R380" ||			// Ericsson R380
      $browser=="UP.B" ||			// UP.Browser
      $browser=="WinW" ||			// WinWAP browser
      $browser=="UPG1" ||			// UP.SDK 4.0
      $browser=="upsi" ||			// another kind of UP.Browser ??
      $browser=="QWAP" ||			// unknown QWAPPER browser
      $browser=="Jigs" ||			// unknown JigSaw browser
      $browser=="Java" ||			// unknown Java based browser
      $browser=="Alca" ||			// unknown Alcatel-BE3 browser (UP based?)
      $browser=="MITS" ||			// unknown Mitsubishi browser
      $browser=="MOT-" ||			// unknown browser (UP based?)
      $browser=="My S" ||           // unknown Ericsson devkit browser ?
      $browser=="WAPJ" ||			// Virtual WAPJAG www.wapjag.de
      $browser=="fetc" ||			// fetchpage.cgi Perl script from www.wapcab.de
      $browser=="ALAV" ||			// yet another unknown UP based browser ?
      $browser=="Wapa")             // another unknown browser (Web based "Wapalyzer"?)
        {
        //////header("Content-type: text/html.xhtml.xml");
		$xmlyes = true;
    }
    else {
      //////header("Content-type: text/html");
?>
<style type="text/css">
<!--
.style2 {font-size: 14px}
-->
</style><head>
<title>Free Ringtones and HTML? And How did you get here?</title>
<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: large;
	color: #993300;
	font-weight: bold;
}
a {
	font-family: Arial, Helvetica, sans-serif;
	font-size: medium;
	color: #990000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #CCCCFF;
}
a:hover {
	text-decoration: none;
	color: #9999FF;
}
a:active {
	text-decoration: none;
	color: #0000FF;
}
-->
</style>

<?php 

    }
  } 

 if ($xmlyes == true) {
 ?>
<?php echo sprintf("%s%sxml version=\"1.0\"? encoding=\"utf\-8\"%s%s","<","?","?",">"); ?> 
<!DOCTYPE html PUBLIC "-//NOKIA//DTD XHTML Mobile +CHTML 1.0//EN" "http://www.nokia.com/dtd/xhtml-mp-chtml.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Free Ringtones? Want a Profile?</title>
<?php } ?></head>

<?php require_once('../Connections/Epwebdev.php'); ?>
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

echo sprintf('%s%sxml version="1.0" encoding="utf-8"%s%s','<','?','?','>');
echo sprintf(' <!DOCTYPE wml PUBLIC %s-//WAPFORUM//DTD WML 1.3//EN%s','"','"');
echo sprintf(' %shttp://www.wapforum.org/DTD/wml13.dtd%s>','"','"');

?>

<style type="text/css">
<!--
.style2 {color: #FF0000}
.style3 {font-family: Arial, Helvetica, sans-serif}
.style4 {color: #FF9933}
-->
</style>


<wml>

 
  <!-- Possible <head> elements here. --> 
   
  <template> 
    <!-- Template implementation here. --> 
    <do type="prev"><prev/></do> 
  </template> 
   
  <card id="card1" title="Project Alpha"> 
     
    <do type="unknown" label="Next">
      <go href="#card2"/>
    </do> 

		<do type="ADSLCheck" label="Check Exchange for ADSL">
		  	  <go href="#ADSLCheck" />
			
			  <setvar name="areacode" value=""/>
	 	      <setvar name="phoneno" value=""/>
		</do> 	  
		<do type="Sysop" label="Sysop Login">
		  	  <go href="#Sysop"  />
		</do> 	  

	  

    <!-- Additional <do> elements here. --> 
     
    <p align="center" class="style3"> 
      <!-- Card implementation here. --> 
      <big><b>Welcome Note</b></big><br/>
	  Welcome to project alpha WAP interface, at the moment this is farely basic WAP interface, from here you can test your clients Line for ADSL Complancy Issues. You can also log into the site using your sysop username and password this will allow you to view statistics of your clients including basic invoice details.<br/><br/>
    </p> 
	<p align="left" class="style3">
	This is a restricted area and is not open to the public. Please be advised that you cannot access any vital details using this interface.<br/><br/>	
	</p>     
	<p align="right" class="style3">
	Copyright 2005 &copy; All Rights Reserved.
</p>     

    <span class="style3">
  <!-- Additional <<p> elements here. --> 
  </card> 
   
  <card id="card2" title="Card #2"> 
    </card>
    </span>
    <card id="card2" title="Card #2"><p align="center" class="style3"> 
      <big><b>
      Second Card</b></big> 
    </p> 
  </card> 
	
	<card id="ADSLCheck" title="ADSL and g.SHDL Availablilty Test"><p align="center" class="style3"><big><b>From here you can <span class="style4">test</span> the line to see if it will support ADSL.</b></big><br/>
	  Simply type in the area code and phone number and the WAP interface will search for the details pertaining to this Number and Exchange.<br/>
	  
 <br/>
	  </p>
	  <form name="form1" method="POST" action="<?php echo $editFormAction; ?>">
	  <fieldset title="Australian Landline Number">
	  <p>
	  <table border="0" align="center">
          <tr>
            <td class="style3">Landline Number: </td>
            <td class="style3"><input name="phoneno" type="text" format="*N" maxlength="8"/></td>
          </tr>
          <tr>
            <td class="style3">Area Code:</td>
            <td class="style3"><input name="areacode" type="text" value="0" maxlength="2" format="*N" /></td>
          </tr>
      </table>
		</p>
		</fieldset>
	</form>
        <p align="center" class="style3">		</fieldset>
		  <span class="style2">Both fields required to see if your exchange supports DSL Services 		
	    </span></p>
    <p align="right" class="style3">
	  <do label="Search for DSL Services" type="unknown">
		  <go href="?areacode=$(areacode)&phoneno=$(phoneno)&MM_insert=DSLLogged#ADSLSubmit"  method="post">
			   <refresh>
				   <postfield name="areacode" value="$(areacode)"/>
				   <postfield name="phoneno" value="$(phoneno)"/>
				   <postfield name="MM_insert" value="DSLLogged"/>
			   </refresh>
			   <?php $nRecID_EnquiryID = 0; ?>
		  </go>
	  </do>
  </p>
</card>


	<card id="ADSLSubmit" title="<?php 
	
	if (empty($phoneno)){
		$showform = true;
		}
	else
	{
	
		$showform = false;
		$number = str_replace(' ', '' ,$phoneno);
		$areac = str_replace(' ', '' ,$areacode);
		$url = sprintf("http://www.comcen.com.au/cgi-bin/checkadslnumber3.cgi?task=find&areacode=%s&number=%s&extra=xml&user=ep.net.au",$areac,$number);
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

			if (!$nRecID_EnquiryID <> 0) {
				if ($MM_insert = "DSLLogged") {
					  $insertSQL = sprintf("INSERT INTO dslcheck (areacode, phone, state, status, phonenumber, exchange) VALUES (%s, %s, %s, %s, %s, %s)",
										   GetSQLValueString($areacode, "text"),
										   GetSQLValueString($phoneno, "text"), GetSQLValueString($state, "text"), GetSQLValueString($status, "text"), GetSQLValueString($number, "text"), GetSQLValueString($exchange, "text"));
					
					  mysql_select_db($database_Epwebdev, $Epwebdev);
					  $Result1 = mysql_query($insertSQL, $Epwebdev) or die(mysql_error());
					
					  $nRecID_EnquiryID = mysql_insert_id();
					}
				}
			
			if ($nRecID_EnquiryID <> 0) {
				mysql_select_db($database_Epwebdev, $Epwebdev);
				$query_EnquiryID = sprintf("SELECT dslcheck.RecID, dslcheck.`state`, dslcheck.queried FROM dslcheck WHERE dslcheck.RecID = %s", $nRecID_EnquiryID);
				$EnquiryID = mysql_query($query_EnquiryID, $Epwebdev) or die(mysql_error());
				$row_EnquiryID = mysql_fetch_assoc($EnquiryID);
				$totalRows_EnquiryID = mysql_num_rows($EnquiryID);
			}

		}
		
	print "$exchange"; ?>"><p align="center" class="style3">
    <big><b>The Following Result was returned by the server.</b></big> <br/> <br/>    
	<br/> 
		The phone line indeed support DSL services. This means that this line can be connected to DSL through your vendor.<br/>
		Make sure you check with the customer/client that they do not already have DSL services to this line otherwise if the do <br/>
		suggest you Churn the over immediately. <br/><br/>
	  <table width="90%" border="0" align="center">
		<tr>
		  <td width="74" class="style3"><div align="right">Enquiry ID:</div></td>
		  <td width="41" class="style3">&nbsp;</td>
		  <td width="272" class="style3"><div align="center"><span class="style8"><?php echo $row_EnquiryID['state']; ?><?php echo $row_EnquiryID['RecID']; ?></span></div></td>
		</tr>
		<tr>
		  <td class="style3"><div align="right"></div></td>
		  <td class="style3">&nbsp;</td>
		  <td class="style3"><div align="center"></div></td>
		</tr>
		<tr>
		  <td class="style3"><div align="right">Status:</div></td>
		  <td class="style3">&nbsp;</td>
		  <td class="style3"><div align="center"><?php print "$status"; ?>
		  </div></td>
		</tr>
		<tr>
		  <td class="style3"><div align="right">Number:</div></td>
		  <td class="style3">&nbsp;</td>
		  <td class="style3"><div align="center"><?php print "$number"; ?>
			</div></td>
		</tr>
		<tr>
		  <td class="style3"><div align="right">Exchange:</div></td>
		  <td class="style3">&nbsp;</td>
		  <td class="style3"><div align="center"><?php print "$exchange"; ?>
	</div></td>
		</tr>
		<tr>
		  <td class="style3"><div align="right">State:</div></td>
		  <td class="style3">&nbsp;</td>
		  <td class="style3"><div align="center"><?php print "$state"; ?>
	</div></td>
		</tr>
	</table>
			<span class="style3">
			
		<?php
			if ($totalRows_EnquiryID > 0) {
				mysql_free_result($EnquiryID);
			} else {
		?> <br/>No result found for ($(areacode)) - $(phoneno) <?php } ?>
    </p> 
        </span>
    </card> 

    <span class="style3">
    <card id="Sysop" title="Username and Password"> 
    </card>
    </span>
    <card id="Sysop" title="Username and Password"><p align="center" class="style3"> 
      <big><b>Please enter in your sysop username and password.</b></big> <br/> <br/>
	  
    </p> 
  </card> 
 
</wml> 