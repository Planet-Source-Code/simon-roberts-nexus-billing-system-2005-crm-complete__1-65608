<?php 
if(!session_id()){
  session_start();
}
require_once('../connections/projectalpha.php');

$hn = gethostbyaddr($REMOTE_ADDR);
if (stristr($hn, "no1.com.au")) {
Header("HTTP/1.1 404 Not Found");
print "<!DOCTYPE HTML PUBLIC \"-//IETF//DTD HTML 2.0//EN\">
	<HTML><HEAD>
	<TITLE>404 Not Found</TITLE>
	</HEAD><BODY>
	<H1>Not Found</H1>
	The requested URL / was not found on this server, Or you do not have access or your domain is banned..<P>
	<P>Additionally, a 404 Not Found
	error was encountered while trying to use an ErrorDocument to handle the request.
	<HR>
	<ADDRESS>Apache/1.3.26 Server at <your_domain> Port 80</ADDRESS>
	</BODY></HTML>";
	exit;
}

function EscapeChars($InputValue) 
	{
	
		$InputValue = str_replace( '\\', '\\\\',$InputValue); 
		
		$InputValue = str_replace(chr(0), '\\0', $InputValue); 
		
		$InputValue = str_replace('\'', '\\\'', $InputValue); 
		$InputValue = str_replace( chr(34) , sprintf("%s%s",'\\' ,chr(34)),  $InputValue);
		$InputValue = str_replace( chr(8), '\\b', $InputValue);
		$InputValue = str_replace( chr(10), '\\n', $InputValue);
		$InputValue = str_replace(chr(13), '\\r', $InputValue);
		$InputValue = str_replace( chr(9), '\\t', $InputValue);
		$InputValue = str_replace( Chr(26), '\\z', $InputValue); 
		return $InputValue;
	}



	if ($logout <> 0) {
	
	
		if ($_SESSION['SysopID'] <> 0) {
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_InCLS1 = sprintf("update sysops set phpsessionid = NOW() where RecID = %d",$_SESSION['SysopID']);
			$InCLS1 = mysql_query($query_InCLS1, $projectalpha) or die(mysql_error());
		}
				
		$_SESSION['SysopID'] = 0;
		$_SESSION['sysCreated'] = 0;
		$_SESSION['StartForcast'] = 0;
		$_SESSION['EndForcast'] = 0;
		$_SESSION['VirtualID'] = 0;
		$_SESSION['Master'] = 0;
		$_SESSION['bMaintain'] = 0;
		$_SESSION['Security'] = 0;
		$_SESSION['bWEBAccount'] = 0;
		$_SESSION['step'] = 0;
		$_SESSION['INClause'] = "";
		$Username = "";
		$Password = "";
		$userlogin = "";
		$logout =0;
		$showtable = true;
		?>
		<HEAD>
		<META HTTP-EQUIV="refresh" CONTENT="0;<?php echo sprintf("URL=http://www.projectalpha.com.au/en/login.php?xSSL='%d'",$xSSL); ?>>
		<style type="text/css">
<!--
.style19 {font-family: "Trebuchet MS", Tahoma, Arial}
.style21 {
	color: #996666;
	font-weight: bold;
	font-size: medium;
}
-->
        </style>
		</HEAD>
		<?php
		exit;
	}
	
require("CustomSql.inc.h4x0r.php");
$db = new CustomSQL($DBName);
$showtable = true;
$errortag = false;
	

$customerid = $db->logincheck('','',session_id());
	
if ($customerid==0 || $_SESSION['SysopID']==0 || empty($_SESSION['INClause']) || $_SESSION['INClause'] == "")  {
	
		if (empty($Username)){
			$forgotusername=true;
			$errortag = true;
			$showtable=true;
			$errormsg = 'You have not entered in a username, please attempt to log on with correct details.';
		}
		
		if (empty($pwd)) {
			if (empty($Password)){
				$forgotpassword=true;
				$errortag = true;
				$showtable=true;
				$errormsg = 'You have not entered in a password, please attempt to log on with correct details.';
			}
				
		}
		if (empty($Password) && empty($Username) && empty($pwd)) {
			$customerid = $db->logincheck('','',session_id());
		} else { 
			if (!empty($pwd)) {
				$customerid = $db->logincheck($Username,$pwd,session_id());
				$errortag = false;
			} else {
				$customerid = $db->logincheck($Username,md5($Password),session_id());
				$errortag = false;
			}
		}
			
		if (!empty($customerid[0]['RecID'])){
			$showtable=false;
			$errortag=false;
			session_cache_expire(60*30);
			$_SESSION['SysopID'] = $customerid[0]['RecID'];
			$_SESSION['sysCreated'] = $customerid[0]['DateCreated'];
			$_SESSION['StartForcast'] = $customerid[0]['StartForcast'];
			$_SESSION['EndForcast'] = $customerid[0]['EndForcast'];
			$_SESSION['Security'] = $customerid[0]['SecurityLevel'];
			$_SESSION['bWEBAccount'] = $customerid[0]['bWEBAccount'];
			$_SESSION['bConfirmed'] = $customerid[0]['bConfirmed'];
			$_SESSION['VirtualID'] = $customerid[0]['VirtualID'];
				mysql_select_db($database_projectalpha, $projectalpha);
				$query_InCLS1 = sprintf("select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID  and  vispb.VirtualID = %d",$_SESSION['VirtualID']);
				$InCLS1 = mysql_query($query_InCLS1, $projectalpha) or die(mysql_error());
				$row_InCLS1 = mysql_fetch_assoc($InCLS1);
				$totalRows_InCLS1 = mysql_num_rows($InCLS1);
				$bStep = $bStep = 1;
				$ib = $ib = " ";
				
				  if ($totalRows_InCLS1 < 1 ) {
					$ib = "(".sprintf("\'%s\'",$_SESSION['VirtualID']).")";
				  }else{
					$ib = "(";
					while ($row_InCLS1 = mysql_fetch_assoc($InCLS1))
						{
							$pos      = strpos($ib, sprintf("\'%s\'",$row_InCLS1['RecIDb']),1);
							if ($pos === false) {
								$ic = ++$bStep;
								$ib .= sprintf("\'%s\', ",$row_InCLS1['RecIDb']);
		
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_InCLS12 = sprintf('select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID and vispb.VirtualID = %d',$row_InCLS1['RecIDb']);
								$InCLS2 = mysql_query($query_InCLS12, $projectalpha) or die(mysql_error());
					
								while ($row_InCLS2 = mysql_fetch_assoc($InCLS2)) 
								{
									$pos      = strpos($ib, sprintf("\'%s\'",$row_InCLS2['RecIDb']),1);
									if ($pos === false) {
				
										$ib .= sprintf("\'%s\', ",$row_InCLS2['RecIDb']);
										$ic = ++$bStep;
		
		
										mysql_select_db($database_projectalpha, $projectalpha);
										$query_InCLS13 = sprintf('select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID and vispb.VirtualID = %d',$row_InCLS2['RecIDb']);
										$InCLS3 = mysql_query($query_InCLS13, $projectalpha) or die(mysql_error());
							
										while ($row_InCLS3 = mysql_fetch_assoc($InCLS3)) 
										{
											$pos      = strpos($ib, sprintf("\'%s\'",$row_InCLS3['RecIDb']),1);
											if ($pos === false) {
						
												$ib .= sprintf("\'%s\', ",$row_InCLS3['RecIDb']);
												$ic = ++$bStep;
		
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_InCLS14 = sprintf('select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID and vispb.VirtualID = %d',$row_InCLS3['RecIDb']);
												$InCLS4 = mysql_query($query_InCLS14, $projectalpha) or die(mysql_error());
									
												while ($row_InCLS4 = mysql_fetch_assoc($InCLS4)) 
												{
													$pos      = strpos($ib, sprintf("\'%s\'",$row_InCLS4['RecIDb']),1);
													if ($pos === false) {
								
														$ib .= sprintf("\'%s\', ",$row_InCLS4['RecIDb']);
														$ic = ++$bStep;
		
														mysql_select_db($database_projectalpha, $projectalpha);
														$query_InCLS15 = sprintf('select distinct vispa.RecID as RecIDa, vispb.RecID as RecIDb, vispb.Description from accountviewer, virtualisp as vispa, virtualisp as vispb where vispa.RecID = vispb.VirtualID and vispb.VirtualID = %d',$row_InCLS3['RecIDb']);
														$InCLS5 = mysql_query($query_InCLS15, $projectalpha) or die(mysql_error());
											
														while ($row_InCLS5 = mysql_fetch_assoc($InCLS5)) 
														{
															$pos      = strpos($ib, sprintf("\'%s\'",$row_InCLS5['RecIDb']),1);
															if ($pos === false) {
										
																$ib .= sprintf("\'%s\', ",$row_InCLS5['RecIDb']);
																$ic = ++$bStep;
															}
														}			
														mysql_free_result($InCLS5);
													}
												}			
												mysql_free_result($InCLS4);
											}
										}			
										mysql_free_result($InCLS3);
		
									}
								}			
								mysql_free_result($InCLS2);
		
							}
		
						}
					mysql_free_result($InCLS1);
		
					$pos = strpos($ib, sprintf("\'%s\'",$_SESSION['VirtualID']),1);
					if ($pos === false) {
		
						$ib .= sprintf("\'%s\')",$_SESSION['VirtualID']);
						$ic = ++$bStep;
						
					} ELSE {
					
						$ib = substr($ib,0,strlen($ib)-2);
						$ib .= ")";
						
					}
		
					$Rst = mysql_query(sprintf('update sysops set INClause=\'%s\' where RecID = \'%d\'',$ib,$_SESSION['SysopID']), $projectalpha) or die(mysql_error());		
					$ic = --$bStep;
					$_SESSION['VISPCount'] = $ic;
					$_SESSION['INClause']  = str_replace('\\\'','\'',$ib);
					
					}
				$showtable = false;
			 } 
	} else {
		$showtable = false;
	}
	if ($_SESSION['INClause']=="") {
		$_SESSION['INClause']  = sprintf("('%d')",$_SESSION['VirtualID']);
	}
	if ($showtable==false) {
		if ($_SESSION['VISPCount']==0) {
			$_SESSION['INClause']  = sprintf("('%d')",$_SESSION['VirtualID']);
			$_SESSION['VISPCount'] = 1;
		}
	}
	if ($errortag==true) {
		$showtable = true;
	}
	if ($showtable ==false) {
		$errortag=false;
		$forgotusername=false;
		$forgotpassword=false;
	}
	if ($showtable == false) {
		if (!empty($Redir)) {
		?>
		<HEAD>
		<META HTTP-EQUIV="refresh" CONTENT="0;URL=<?php echo $Redir; ?>">
		</HEAD>
		<?php
		exit;
		}
	}
	
	
?>



<html>
<head>
<title><?php print "$front_login"; ?></title>
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
<link href="ttxbox.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style3 {font-size: 36px}
.style4 {
	color: #990066;
	font-weight: bold;
}
body {
	background-color: #DEDECA;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
}
.style6 {font-size: 24px}
.style10 {font-size: 18; font-weight: bold; }
.style15 {font-size: medium; color: #0000CC; }
.style20 {font-size: 18; font-weight: bold; font-family: "Trebuchet MS", Tahoma, Arial; }
-->
</style>
</head>

<body text="#000000">
<?php
include("top.php3");
?>
<table width="770" border="0" cellspacing="1" cellpadding="0" align="center" class="table_01">
  <tr> 
    <td width="171" align="center" valign="top" class="table_02"> 
    <table width="160" border="0" cellpadding="4" cellspacing="0" class="table_02">
        <tr> 
          <td><div align="center"><img src="http://www.projectalpha.com.au/images/idents/ep_ident.jpg" width="164" height="170"></div></td>
        </tr>
        <tr> 
          <td><div align="center">ACN 096 867 775 </div></td>
        </tr>
      </table>
   
      <div align="right">
            <p align="center">
			<?php
			if ($xSSL==0) {
			?>
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="147" height="34">
                <param name="movie" value="button1.swf">
                <param name="quality" value="high">
                <param name="bgcolor" value="#F0F0F0">
                <embed src="button1.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="147" height="34" bgcolor="#F0F0F0"></embed>
              </object>
          <?php } ?>
          </p>
      </div>      <div align="center">      
      <p align="center">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="23">
          <param name="movie" value="logout.swf">
          <param name="quality" value="high">
          <param name="bgcolor" value="#DDDDDD">
          <embed src="logout.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="23" bgcolor="#DDDDDD"></embed>
        </object>
</p>
      </div></td>
      <td class="menu" bgcolor="#FFFFFF" valign="top" width="601"> 
     <table border="0" cellspacing="0" cellpadding="4" width="599">
        <tr> 
          <td width="591" bgcolor="#F2F2F2" class="menu_in">::<?php print "$front_login"; ?>&nbsp;&nbsp;&nbsp;<font color="#FF0000">*</font>&nbsp;<?php print "$front_requiredinfo"; ?></td>
        </tr>
        
        <tr> 
          <td>
		  <?php
        if ($forgotusername==true || $showtable==true){
        ?>		<table width="90%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="7%" align="center" valign="top"><img src="images/warning.gif" width="28" height="28"></td>
              <td width="93%"><blockquote class="style20">Missing Valid Username for log-on.</blockquote></td>
            </tr>
          </table>
		  <?php
		  }
        if ($forgotpassword==true || $showtable==true){
        ?>
		  <table width="90%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="7%" align="center" valign="top"><img src="images/warning.gif" width="28" height="28"></td>
              <td width="93%"><blockquote class="style20">Please enter in a valid password. If you have forgotton your password then click on the Green Button titled 'Password???' to retrieve your <em><strong>passwo</strong></em>rd. </blockquote></td>
            </tr>
          </table>          </td>
        </tr>
        <?php
        }
        ?>
        <tr> 
          <td>
            <p>
              <?php
			  //if ((empty($_SESSION['VISPCount'])) || ($_SESSION['VISPCount'] ==0)) {
			  	//$showtable=true;
				//}
            if ($showtable){
            ?>
            </p>
            <form action="<?php print "$PHP_SELF"; ?>" method="POST" name="userlogin" id="userlogin">
<table border=0 cellpadding=2 cellspacing=2>
<tr>
  <td width="120" class="style3"><span class="style19 style15"><strong>Username:</strong></span></td>
  <td width="455"><div align="center">
    <input name="Username" type="text" class="style3" value="<?php print "$Username"; ?>" size="18">    
    <span class="style3">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></div></td>
</tr>
<tr>
  <td class="style3"><span class="style19 style15"><strong>Password:</strong></span></td>
  <td><div align="center">
    <input name="Password" type="Password" class="style3" value="" size="18">
      <span class="style3">&nbsp;<font color="#FF0000">*</font><font color="#FF0000"></font></span></div></td>
</tr>
<tr>
  <td class="style3">
    <div align="center" class="style19 style10"><strong>
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="23">
            <param name="BGCOLOR" value="">
            <param name="movie" value="register.swf">
            <param name="quality" value="high">
            <embed src="register.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="23" ></embed>
      </object>
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="23">
            <param name="movie" value="ForgotPassword.swf">
            <param name="quality" value="high">
            <embed src="ForgotPassword.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="23" ></embed>
      </object>
    </strong></div></td>
  <td>
  <div align="center">
    <input name="userlogin" type="submit" class="style6" value="Log Into The Nexus">
  </div></td></tr>
</table>
<?php
	}
	else{
	?>
            </form>
	<table width="98%"  border="0">
      <tr>
        <td width="1%">&nbsp;</td>
        <td colspan="4" rowspan="2"><div align="justify">
          <p align="center">
            <?php
			if ($_SESSION['bConfirmed']==-1 || $_SESSION['bConfirmed']==-3) {
				include('wizard1.php');
				if ($_SESSION['bConfirmed']==-3) {
					mysql_select_db($database_projectalpha, $projectalpha);
					$query_RegStep = sprintf("select RegStep from virtualisp_extended where VirtualID = %d",$_SESSION['VirtualID']);
					$RegStep = mysql_query($query_RegStep, $projectalpha) or die(mysql_error());
					$row_RegStep = mysql_fetch_assoc($RegStep);
					$totalRows_RegStep = mysql_num_rows($RegStep);

					$_SESSION['step']=$row_RegStep['RegStep'];
				}
			} else {
				include('dlwizard1.php');
			}
			?>
</p>
          </div>          <div align="justify">        </div>        <div align="center"></div>        <div align="center"></div></td>
        </tr>
      <tr>
        <td height="16">&nbsp;</td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td colspan="4"><div align="center"></div></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td width="26%"><div align="center"><a href="/en/modiinfo.php"><img src="/en/images/modinfo.png" width="175" height="50" border="0"></a></div></td>
        <td width="24%">&nbsp;</td>
        <td width="23%">&nbsp;</td>
        <td width="26%"><a href="/en/modipass.php"><img src="/en/images/modpass.png" width="175" height="50" border="0"></a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><span class="style21">Primary Sysop Functions </span></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><a href="publishstationary.php"><strong>Publish Stationary </strong></a></div></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><a href="publishnews.php"><strong>Publish News </strong></a></div></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td colspan="5"><h3 class="style4">Sysop  Quick Functions </h3></td>
        </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><a href="clnt_explorer.php"><strong>View Your Clients</strong></a> </div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"><a href="../cwan/maintenance/doinvoices.php"><strong>Process and Email Invoices </strong></a> </div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><a href="inv_explorer.php"><strong>View Your Active Invoices</strong></a> </div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><a href="sales.php"><strong>View Your <em>Hot Link</em> Sales Report</strong></a> </div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><strong><a href="dateinterval.php">Set Date Interval On Sales Report </a></strong></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>Number of Resellers, Wholsalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP's Mapped: </td>
        <td colspan="3"><?php echo $_SESSION['VISPCount'] ?></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>Resellers, Wholsalers, Telco &amp; Telecom Enterprise, ISP &amp; ViSP IN Clause: </td>
        <td colspan="3"><?php echo $_SESSION['INClause'] ?></td>
        </tr>
    </table>	
	<?php
	}
	?>          </td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
        </tr>
      </table>    </td>
  </tr>
</table>
<?php
include("bottom.php3");
?>
</body>
</html>

