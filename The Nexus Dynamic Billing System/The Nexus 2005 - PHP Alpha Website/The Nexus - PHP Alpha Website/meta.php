<?php require_once('Connections/Epwebdev.php'); ?>
<?php
if ($groupid == 0) {
	$groupid = 1;
}

mysql_select_db($database_Epwebdev, $Epwebdev);
$query_meta = "SELECT `_nexus_metasystem`.`mt_name_user-defined`, `_nexus_metasystem`.`mt_http-equiv_user-defined`, `_nexus_metasystem`.mt_name, `_nexus_metasystem`.`mt_http-equiv`, `_nexus_metasystem`.mt_Content FROM `_nexus_metasystem` WHERE `_nexus_metasystem`.hl_type = 'META' and `_nexus_metasystem`.hl_groupid = $groupid";
$meta = mysql_query($query_meta, $Epwebdev) or die(mysql_error());
$totalRows_meta = mysql_num_rows($meta);
if ($totalRows_meta > 0) {
	while ($row_meta = mysql_fetch_assoc($meta))
	{
		if ($row_meta['mt_name'] == 'user-defined') {
			$mtname = $row_meta['mt_name_user-defined'];
		} else {
			$mtname = $row_meta['mt_name'];
		}
		if ($row_meta['mt_http-equiv'] == 'user-defined') {
			$mthttp = $row_meta['mt_http-equiv-defined'];
		} else {
			$mthttp = $row_meta['mt_http-equiv'];
		}
		
		$rowmeta = $row_meta['mt_Content'];
		if (strlen($nexpv01) > 0) {
			$rowmeta = str_replace('\\\'','',str_replace("<NEXPV01 />", $nexpv01, $rowmeta));
		} else {
			$rowmeta = str_replace("<NEXPV01 />,", '', $rowmeta);
		}
		
		if (strlen($nexpv02) > 0) {
			$rowmeta = str_replace('\\\'','',str_replace("<NEXPV02 />", $nexpv02 . ' - ', $rowmeta));
		} else {
			$rowmeta = str_replace("<NEXPV02 />", '', $rowmeta);
		}
		
		print sprintf("<meta name=\"%s\" http-equiv=\"%s\" content=\"%s\">%s", $mtname, $mthttp, $rowmeta,chr(13));
	}
}
mysql_free_result($meta);

mysql_select_db($database_Epwebdev, $Epwebdev);
$query_linker = "SELECT `_nexus_metasystem`.lk_rel, `_nexus_metasystem`.lk_href, `_nexus_metasystem`.lk_charset, `_nexus_metasystem`.lk_class, `_nexus_metasystem`.lk_dir, `_nexus_metasystem`.lk_hreflang, `_nexus_metasystem`.lk_type, `_nexus_metasystem`.lk_title, `_nexus_metasystem`.lk_label, `_nexus_metasystem`.lk_id FROM `_nexus_metasystem` WHERE `_nexus_metasystem`.hl_type = 'LINK' AND `_nexus_metasystem`.hl_groupid = $groupid";
$linker = mysql_query($query_linker, $Epwebdev) or die(mysql_error());
$totalRows_linker = mysql_num_rows($linker);
if ($totalRows_linker > 0) {
	while ($row_linker = mysql_fetch_assoc($linker))
	{
		$lktxt = "";
		
		if (!empty($row_linker['lk_rel'])) {
			$lktxt .= " rel=\"" . $row_linker['lk_rel'] . "\"";
		}
		if (!empty($row_linker['lk_href'])) {
			$lktxt .=  " href=\"" . $row_linker['lk_href'] . "\"";
		}
		if (!empty($row_linker['lk_charset'])) {
			$lktxt .=  " charset=\"" . $row_linker['lk_charset'] . "\"";
		}
		
		if (!empty($row_linker['lk_class'])) {
			$lktxt .=  " class=\"" . $row_linker['lk_class'] . "\"";
		}
		
		if (!empty($row_linker['lk_dir'])) {
			$lktxt .=  " dir=\"" . $row_linker['lk_dir'] . "\"";
		}
		
		if (!empty($row_linker['lk_hreflang'])) {
			$lktxt .=  " hreflang=\"" . $row_linker['lk_hreflang'] . "\"";
		}
		
		if (!empty($row_linker['lk_type'])) {
			$lktxt .=  " type=\"" . $row_linker['lk_type'] . "\"";

		}
		
		if (!empty($row_linker['lk_title '])) {
			$lktxt .=  " title =\"" . $row_linker['lk_title'] . "\"";

		}
		
		if (!empty($row_linker['lk_lang '])) {
			$lktxt .=  " lang=\"" . $row_linker['lk_lang'] . "\"";
		}
		
		if (!empty($row_linker['lk_label '])) {
			$lktxt .=  " label =\"" . $row_linker['lk_label'] . "\"";
		}
		
		if (!empty($row_linker['lk_id '])) {
			$lktxt .=  " title =\"" . $row_linker['lk_id'] . "\"";
		}
		print sprintf("<link%s>%s", $lktxt,chr(13));
	}
}




mysql_free_result($linker);
?>
