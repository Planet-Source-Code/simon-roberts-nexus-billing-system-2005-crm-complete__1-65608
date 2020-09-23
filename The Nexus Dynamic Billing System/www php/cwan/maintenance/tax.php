<?php require_once('../../Connections/projectalpha.php'); 


function vidtax($VirtualID) {

	mysql_select_db($database_projectalpha, $projectalpha);
	$query_sqlz = sprintf("select cTaxCode, cTaxCountry, bTaxMode, cTaxExemptCode Where RecID = '%d%'",$row_cyclemet['RadiusID']));
	$vidtax1 = mysql_query($query_sqlz, $projectalpha) or die(mysql_error());
	$row_vidtax1 = mysql_fetch_assoc($vidtax1);
	$totalRows_vidtax1 = mysql_num_rows($vidtax1);

	mysql_select_db($database_projectalpha, $projectalpha);
	$query_sqlx = sprintf("select Percentage Where Code = '%s%' and Country = '%s'",$row_vidtax1['cTaxCode'],$row_vidtax1['cTaxCountry']));
	$vidtax2 = mysql_query($query_sqlx, $projectalpha) or die(mysql_error());
	$row_vidtax2 = mysql_fetch_assoc($vidtax2);
	$totalRows_vidtax1 = mysql_num_rows($vidtax2);

	return $row_vidtax2['Percentage'];

}

?>