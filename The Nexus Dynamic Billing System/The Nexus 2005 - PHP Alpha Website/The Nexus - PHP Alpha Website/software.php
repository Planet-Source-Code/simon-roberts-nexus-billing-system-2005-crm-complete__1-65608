<?php require_once('Connections/Epwebdev.php'); ?>

<?php

if (empty($SoftLabCat)) {
	$SoftLabCat = 'Analyst';
	$SoftLabID = 1;
}

if (empty($SoftLabID)) {
	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_Recordset1 = "SELECT software.ID, software.Category, software.SubCategoryOf, software.Title, software.SubTitle, software.ColumnA, software.ColumnB, software.Version, software.Author, software.Licence, software.DownloadCRC, software.DownloadCRCType, software.DownloadURL, software.SourceURL, software.WWWURL, software.ScreenShotURL, software.AppendixImageURL, software.BottomIncludeURL, software.DownloadFilename, software.AppendixHTML FROM software WHERE Category = '$SoftLabCat'";
	$Recordset1 = mysql_query($query_Recordset1, $Epwebdev) or die(mysql_error());
	$row_Recordset1 = mysql_fetch_assoc($Recordset1);
	$totalRows_Recordset1 = mysql_num_rows($Recordset1);
	$SoftLabID = $row_Recordset1['ID'];
} else {
	mysql_select_db($database_Epwebdev, $Epwebdev);
$query_Recordset1 = "SELECT software.ID, software.Category, software.SubCategoryOf, software.Title, software.SubTitle, software.ColumnA, software.ColumnB, software.Version, software.Author, software.Licence, software.DownloadCRC, software.DownloadCRCType, software.DownloadURL, software.SourceURL, software.WWWURL, software.ScreenShotURL, software.AppendixImageURL, software.BottomIncludeURL, software.DownloadFilename, software.AppendixHTML FROM software WHERE ID = '$SoftLabID'";
$Recordset1 = mysql_query($query_Recordset1, $Epwebdev) or die(mysql_error());
$row_Recordset1 = mysql_fetch_assoc($Recordset1);
$totalRows_Recordset1 = mysql_num_rows($Recordset1);
}

	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_Recordset3 = "SELECT software.ID FROM software WHERE Category = '$SoftLabCat'";
	$Recordset3 = mysql_query($query_Recordset3, $Epwebdev) or die(mysql_error());
	$totalRows_Recordset3 = mysql_num_rows($Recordset3);

	if ($totalRows_Recordset3 == 1) {
		$NoDir =true;
	} else {
		$PrevSet = false;
		while ($row_Recordset3 = mysql_fetch_assoc($Recordset3)) {
			
			if ($row_Recordset3['ID'] == $SoftLabID) {
				if (!empty($PrevID)) {
					$PrevSet=true;
				}
				$NextSet=false;
				while ($row_Recordset3 = mysql_fetch_assoc($Recordset3)) {
					if ($NextSet==false) {
						$NextID=$row_Recordset3['ID'];
						$NextSet=true;
					}
				}
			}
			if ($PrevSet == false) {
				$PrevID=$row_Recordset3['ID'];
			}
			
		}

	}




?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Software Catagory: <?php echo $row_Recordset1['Category']; ?> - <?php echo $row_Recordset1['Title']; ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
a {
	font-family: Trebuchet MS, Tahoma, Arial;
	color: #00CCFF;
}
a:link {
	text-decoration: none;
	color: #000000;
}
a:visited {
	text-decoration: none;
	color: #993333;
}
a:hover {
	text-decoration: none;
	color: #CC0000;
}
a:active {
	text-decoration: none;
	color: #FF6633;
}
.style3 {
	font-family: "Trebuchet MS", Tahoma, Arial;
	font-style: italic;
	font-weight: bold;
	font-size: 16px;
	color: #666666;
}
.style6 {color: #FFFFFF; font-weight: bold; }
.style7 {color: #FFFFFF}
.style13 {
	color: #FFCC66;
	font-weight: bold;
	font-size: small;
}
.style15 {
	font-size: small;
	font-weight: bold;
	color: #9933CC;
}
.style16 {color: #000000}
.style17 {font-family: "Trebuchet MS", Tahoma, Arial}
.style18 {font-size: x-small; font-family: "Trebuchet MS", Tahoma, Arial; }
.style19 {
	color: #333333;
	font-size: medium;
}
body,td,th {
	color: #000000;
}
body {
	background-color: #000000;
}
.style20 {font-size: small}
.style21 {
	font-size: 24px;
	font-weight: bold;
}
-->
</style>
</head>

<body>
<table width="95%"  border="0" align="center" bgcolor="#FFCC99">
  <tr bgcolor="#66CC99">
    <td colspan="2"><table width="100%"  border="0" align="center" cellpadding="1" cellspacing="0">
      <tr>
        <td width="13%" align="left" class="style18"><img src="images/icons/softlab_96x96.gif" width="96" height="110"></td>
        <td width="63%" class="style18"><div align="center" class="style15 style19"><span class="style21"><?php echo $row_Recordset1['Title']; ?></span><br>
          Software Catagory: <?php echo $row_Recordset1['Category']; ?></div></td>
        <td width="24%" align="right" valign="top" class="style18"><p>
          <?php 
			if ($PrevSet==true || !$PrevID=='') { ?>
          <a href="software.php?SoftLabID=<?php echo $PrevID ?>&SoftLabCat=<?php echo $row_Recordset1['Category']; ?>"><img src="images/icons/LastPage.gif" width="155" height="35" border="0"></a>
          <?php } ?>
          <br>
          <?php 
			if ($NextSet==true) { ?>
      <a href="software.php?SoftLabID=<?php echo $NextID ?>&SoftLabCat=<?php echo $row_Recordset1['Category']; ?>"><img src="images/icons/NextPage.gif" width="155" height="34" border="0"> </a>
                  <?php } ?>
          </p></td>
      </tr>
    </table></td>
    <td width="18%" rowspan="10" align="center" valign="top" bgcolor="#66CCCC"><table width="86%" height="281"  border="0" cellpadding="2" cellspacing="1" bordercolor="#000000">
      <tr align="center">
        <th colspan="2" nowrap bgcolor="#009999" class="style18"><span class="style13">Software Categories </span></th>
        </tr>
	<?php
	mysql_select_db($database_Epwebdev, $Epwebdev);
	$query_Cats = "select distinct software.Category FROM software";
	$Cats = mysql_query($query_Cats, $Epwebdev) or die(mysql_error());
	while ($row_Cats = mysql_fetch_assoc($Cats)) { ?>
      <tr bgcolor="#66CCCC">
        <td width="10" height="22" nowrap class="style18"><img src="en/images/bullet_b.gif" width="10" height="17"></td>
        <td width="125" nowrap bgcolor="#66CCCC" class="style18"><a href="software.php?SoftLabCat=<?php echo $row_Cats['Category']; ?>" class="style16 style20"><strong><?php echo $row_Cats['Category']; ?></strong></a></td>
      </tr><?php } ?>
      <tr bgcolor="#66CCCC">
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr bgcolor="#66CCCC">
        <td colspan="2" bgcolor="#009999"> <IFRAME ID=IFrame1 FRAMEBORDER=0 SCROLLING=NO 
SRC="http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1" 
height=260 width=170 bgcolor=009999>
Your browser does not support inline frames...However, you can click 
<A 
href="http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=1">
here</a> to see the related document.
</IFRAME></td>
        </tr>
	
    </table></td>
  </tr>
  <tr>
    <td colspan="2"><h1 align="center" class="style17"><span class="style3"><?php echo $row_Recordset1['SubTitle']; ?></span></h1></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center"><span class="style17">
      <?php
	if (!empty( $row_Recordset1['ScreenShotURL'])) {
		?>
      <img src="<?php echo $row_Recordset1['ScreenShotURL']; ?>">
      <?php } ?>
    </span></td>
  </tr>
  <tr>
    <td height="60" colspan="2"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="50%" class="style17"><blockquote>
          <p align="justify"><?php echo $row_Recordset1['ColumnA']; ?></p>
        </blockquote></td>
        <td width="50%" class="style17"><div align="justify">
          <blockquote>
            <p><?php echo $row_Recordset1['ColumnB']; ?></p>
          </blockquote>
        </div></td>
      </tr>
    </table>      </td>
  </tr>
  <tr bgcolor="#FFCC99">
    <td width="43%"><span class="style17">
      <?php
	if (!empty( $row_Recordset1['AppendixImageURL'])) {
		?>
        </span>
      <div align="center" class="style17"><img src="<?php echo $row_Recordset1['AppendixImageURL']; ?>"></div>
      <span class="style17">
      <?php } ?>    
    </span></td>
    <td width="39%"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td width="32%" align="right" class="style17"><span class="style6 style16">Version:</span></td>
        <td width="63%" bgcolor="#CCCCCC" class="style17"><div align="center"><strong><?php echo $row_Recordset1['Version']; ?></strong></div></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style6 style16">Author:</span></td>
        <td bgcolor="#CCCCCC" class="style17"><div align="center"><?php echo $row_Recordset1['Author']; ?></div></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style7"><span class="style6 style16">Download <?php echo $row_Recordset1['DownloadCRCType']; ?> CRC :</span></span></td>
        <td bgcolor="#CCCCCC" class="style17"><div align="center"><?php echo $row_Recordset1['DownloadCRC']; ?></div></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style7"><span class="style6 style16">Licence:</span></span></td>
        <td bgcolor="#CCCCCC" class="style17"><div align="center"><?php echo $row_Recordset1['Licence']; ?></div></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style7"><span class="style6 style16">Website:</span></span></td>
        <td bgcolor="#CCCCCC" class="style17"><?php
	if (!empty($row_Recordset1['WWWURL'])) {
		?>		<div align="center"><a href="<?php echo $row_Recordset1['WWWURL']; ?>">(Click here to view)</a></div>
        <?php } ?></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style7"><span class="style6 style16">Installer Download :</span></span></td>
        <td bgcolor="#CCCCCC" class="style17"><?php
	    if (!empty($row_Recordset1['DownloadURL'])) {
		?>		<div align="center"><a href="<?php echo $row_Recordset1['DownloadURL']; ?>"><?php echo $row_Recordset1['DownloadFilename']; ?></a></div>
        <?php } ?></td>
      </tr>
      <tr>
        <td width="32%" align="right" class="style17"><span class="style7"><span class="style6 style16">Source Code:</span></span></td>
        <td bgcolor="#CCCCCC" class="style17"><?php
	    if (!empty($row_Recordset1['SourceURL'])) {
		?>		<div align="center"><a href="<?php echo $row_Recordset1['SourceURL']; ?>">(Click to get the source) </a></div>
        <?php } ?></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle"> </td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle"><?php
	if (!empty($row_Recordset1['AppendixHTML'])) {
		?><?php echo $row_Recordset1['AppendixHTML']; ?>
        <?php } ?></td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle"><h2 align="center" class="style17"></h2></td>
  </tr>
</table>
</body>
</html>
<?php
mysql_free_result($Recordset1);

mysql_free_result($Cats);
?>
