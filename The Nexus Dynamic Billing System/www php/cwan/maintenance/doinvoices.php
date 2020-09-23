<?php require_once('../../Connections/Epwebdev.php'); ?>
<?php require_once('../../Connections/projectalpha.php'); ?>
<?php require_once('tax.php'); ?>
<?php require_once('dates.php'); ?>
<?php require_once('recieptmaker.php'); ?>
<?php require_once('xoopsmailer.php'); ?>

<?php

if(!session_id()){
  session_start();
}

if (!(session_is_registered("SysopID"))){
	print "<a href=\"http://www.projectalpha.com.au/en/login.php\">Please login to Run Maitenance.</a>";
	exit;
}

$lBytesPerMB = 1024 ^ 2;

	if (!empty($cmdKillAllNow)) {
	
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_UpdateCycle = "delete from online_maintenancelocker where Active <> 1";
		$update = mysql_query($query_UpdateCycle, $projectalpha) or die(mysql_error());	
	}

	//1.1
	if (!empty($cmdRunNow)) {
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsVISPb = sprintf("select BigIntSet from online_settingsandtimers where ConfigKey = 'LOCALSERVERID'",$_SESSION['VirtualID']);
		$ServerID = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
		$row_ServerID = mysql_fetch_assoc($ServerID);
		$totalRows_ServerID = mysql_num_rows($ServerID);
		//1.2
		if ($totalRows_ServerID == 0) {
			mysql_select_db($database_Epwebdev, $Epwebdev);
			$query_LongandLatit = sprintf("SELECT map_paservers.Longitude, map_paservers.Latitude, map_paservers.RegionID, map_paservers.CountryID FROM map_paservers where ID = '%d'",1);
			$LongandLatit = mysql_query($query_LongandLatit, $Epwebdev) or die(mysql_error());
			$row_LongandLatit = mysql_fetch_assoc($LongandLatit);
			$totalRows_LongandLatit = mysql_num_rows($LongandLatit);
		}else {
			mysql_select_db($database_Epwebdev, $Epwebdev);
			$query_LongandLatit = sprintf("SELECT map_paservers.Longitude, map_paservers.Latitude, map_paservers.RegionID, map_paservers.CountryID FROM map_paservers where ID = '%d'",$row_ServerID['BigIntSet']);
			$LongandLatit = mysql_query($query_LongandLatit, $Epwebdev) or die(mysql_error());
			$row_LongandLatit = mysql_fetch_assoc($LongandLatit);
			$totalRows_LongandLatit = mysql_num_rows($LongandLatit);
		}
		
		
		mysql_select_db($database_projectalpha, $projectalpha);
		$query_rsVISPb = sprintf("select IntegerSet from online_settingsandtimers where ConfigKey = 'DOINVOICEOFFSET' and VirtualID = '%d'",$_SESSION['VirtualID']);
		$offset = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
		$row_offset = mysql_fetch_assoc($offset);
		$totalRows_offset = mysql_num_rows($offset);
		//1.2
		if ($totalRows_offset > 0) {
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_rsVISPb = sprintf("select NOW() as DateNava, dateadd(now(), \"%d seconds\") as DateNavb",$row_offset['IntegerSet']);
			$datenav = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
			$row_datenav = mysql_fetch_assoc($datenav);
		} else {
			mysql_select_db($database_projectalpha, $projectalpha);
			$query_rsVISPb = sprintf("select NOW() as DateNava, dateadd(now(), \"%d seconds\") as DateNavb",30);
			$datenav = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
			$row_datenav = mysql_fetch_assoc($datenav);
		}
		$KeepLooping =true;
		//1.2
		while ($KeepLooping =true) {
			$bNewGroup = true;
		
			mysql_select_db($database_projectalpha, $projectalpha);
			$sqla = sprintf("insert into online_maintenancelocker (SysopHostName, PHPSESSION, DateNava, DateNavb, SysopIDRunningMaintenance, SvrRecID, SysopID, VirtualID, Active) select '%s', '%s', '%s', '%s', '%d', RecID, SysopID, VirtualID, VirtualID in (%s) from acci_services inner join accountinfo on acci_services.acci_RecID = accountinfo.RecID inner join servicetypes on acci_services.ServiceID = servicetypes.RecID outer right join online_maintenancelocker on online_maintenancelocker.SvrRecID = acci_services.RecID and online_maintenancelocker.Active <> 1 where accountinfo.Cancelled = '0' and NextCycle ='%s' or NextCycle < '%s' order by SubRecID",gethostbyaddr($REMOTE_ADDR), session_id(), $row_datenav['DateNavb'],$row_datenav['DateNava'],$_SESSION['SysopID'],$_SESSION['INClause'],$row_datenav['DateNavb'],$row_datenav['DateNava']);
			$execsqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
			
			mysql_select_db($database_projectalpha, $projectalpha);
			$sqlz="select RecID, ServiceKey, VirtualID, SubRecID, ptRecID, acci_RecID, RadiusID, SysopID, NextCycle, PreviousCycle, Activation, JoiningFee, PeriodFee, PerMB, PerHour, crsLocal, crsNational, crsInternational, crsMobile, crsVOIP, crsSpecial, crsFlatRate, crsRateMeasured, crsCallCredits, creLocal, creNational, creInternational, creMobile, creVOIP, creSpecial, creFlatRate, creRateMeasured, creRateMeasuredphp, CallCredits, LineRental, EquipmentHire, ExtendedServiceCost, CallCapping, FlagFall from acci_services inner join servicetypes on acci_services.ServiceID = servicetypes.RecID inner join online_maintenancelocker on online_maintenancelocker.SysopID = acci_services.SysopID and online_maintenancelocker.VirtualID = acci_services.VirtualID and online_maintenancelocker.SvrRecID = acci_service.RecID where online_maintenancelocker.Active = 1 and online_maintenancelocker.SysopIDRunningMaintenance = '%d' and online_maintenancelocker.SysopHostName = '%s' and online_maintenancelocker.PHPSESSION = '%s'";
			$query_rsVISPb = sprintf($sqlz,$_SESSION['SysopID'],gethostbyaddr($REMOTE_ADDR), session_id());
			$cyclemet = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
			$totalRows_cyclemet = mysql_num_rows($rsVISPb);
			//1.3
			if ($totalRows_cyclemet==0) {
				$KeepLooping = false;
			}
			//1.3
			if ($KeepLooping =true) {
				//1.4
				while ($row_cyclemet = mysql_fetch_assoc($cyclemet)) 
				{
						//1.5
						if ($row_cyclemet['ServiceKey'] == 'VOIP') {
							// VOIP BILLING GOES HERE
						} else {
							if ($tsubrecid <> $row_cyclemet['SubRecID']) {
								$bNewGroup = true;
								$tsubrecid = $row_cyclemet['SubRecID'];
							}
							//1.6
							if ($row_cyclemet['RadiusID'] <> 0) 
								{
				
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_sqla = sprintf("select RecID, sfCycle_Upload , sfCycle_Download, sfCycle_Mins, VirtualID from radiusaccounts Where RecID = '%d%'",$row_cyclemet['RadiusID']);
								$rsRadius = mysql_query($query_sqla, $projectalpha) or die(mysql_error());
								$row_rsRadius = mysql_fetch_assoc($rsRadius);
								$totalRows_cyclemet = mysql_num_rows($rsRadius);
								
								$cCharge = 0;
				
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_sqlb = sprintf("select * from plantypes where RecID ='%s'",$row_cyclemet['ptRecID']);
								$plantype = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
								$totalRows_plantype = mysql_num_rows($plantype);
								$row_plantype = mysql_fetch_assoc($plantype);
								//1.7
								if ($totalRows_plantype > 0)
									{
								 
										mysql_select_db($database_projectalpha, $projectalpha);
										$query_sqlc = sprintf("select AccountActive from acci_dslconnections where acci_RecID = '%s'",$row_cyclemet['acci_RecID']);
										$accactive = mysql_query($query_sqlc, $projectalpha) or die(mysql_error());
										$totalRows_accactive = mysql_num_rows($accactive);
										$row_accactive = mysql_fetch_assoc($accactive);
										//1.8
										if ($totalRows_accactive > 0) 
										{
											//1.9
											if ($row_plantype['BillOnce'] ==1 || $row_cyclemet['Activation'] == $row_cyclemet['NextCycle']) 
											{
				
											//mysql_select_db($database_projectalpha, $projectalpha);
											//$query_sqld = sprintf("select * from accountinfo Where RecID = ",$row_cyclemet['acci_RecID']);
											//$accinfo = mysql_query($query_sqld, $projectalpha) or die(mysql_error());
											//$totalRows_accinfo = mysql_num_rows($accinfo);
											//$row_accinfo = mysql_fetch_assoc($accinfo);
				
											$invout_Description = $plantype['Description'];
				
											//1.9
											if ($cyclemet['Activation'] == $row_cyclemet['NextCycle']) 
												{ 
				
													$invout_Description .= sprintf(" - Account Setup Fee [%.2f",$row_cyclemet['JoiningFee']);
													$cCharge = $cCharge + $row_cyclemet['JoiningFee'];
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_rsVISPb = sprintf("Update acci_services set Checked='-1' where RecID = ",$row_cyclemet['RecID']);
													$update = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
				
												}
													
				
											//1.9
											if ($row_cyclemet['MBPerPeriod'] <> -1) 
												{
												//1.10
												if ( $row_rsRadius['sfCycle_Download'] / $lBytesPerMB > $row_plantype['MBPerPeriod']) {
													$invout_Description .= sprintf(" %.2f MB's Over", ($row_rsRadius['sfCycle_Download'] / $lBytesPerMB) -  $row_plantype['MBPerPeriod']);
													$cCharge = $cCharge + (($row_rsRadius['sfCycle_Download'] / $lBytesPerMB -  $row_plantype['MBPerPeriod']) /  $row_plantype['MBBlockSize']) * $row_cyclemet['PerMB'];
												}
											
												$invout_sfCycle_Download = $row_rsRadius['sfCycle_Download'];
												$invout_sfCycle_Upload = $row_rsRadius['sfCycle_Upload'];
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlf = sprintf("select RecID, sfCycle_Upload, sfCycle_Download, sfCycle_Mins from accountinfo Where RecID = '%s'",$row_cyclemet['acci_RecID']);
												$rsAccInfo = mysql_query($query_sqlf, $projectalpha) or die(mysql_error());
												$totalRows_rsAccInfo = mysql_num_rows($accactive);
												$row_rsAccInfo = mysql_fetch_assoc($accactive);
												//1.10
												if ($totalRows_rsAccInfo > 0)
												{
							
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Insert Into history_acci_datausage (acci_RecID, Uploaded, Downloaded, NumMins) VALUES('%s','%d','%d','%d')",$row_cyclemet['acci_RecID'],$row_rsAccInfo['sfCycle_Upload'],$row_rsAccInfo['sfCycle_Download'],$row_rsAccInfo['sfCycle_Mins']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Upload = sfCycle_Upload - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Upload'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Download = sfCycle_Download - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Download'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
										
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Upload = sfCycle_Upload - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Mins'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												   
												}
							
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Download = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Upload = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												//1.10
												if ($row_plantype['HoursPerPeriod'] <> -1)
												{
													//1.11
													if ($row_rsRadius['sfCycle_Mins'] / 60 > $row_plantype['HoursPerPeriod'])
													{
														$invout_Description = sprintf("%s %d Min's Over",$invout_Description, ($row_rsRadius['sfCycle_Mins'] - $row_plantype['HoursPerPeriod'] * 60));
														$cCharge = $cCharge + ($row_rsRadius['sfCycle_Mins'] / 60) * $row_cyclemet['PerHour'];
													}//1.11
												}//1.10
												
												$invout_sfCycle_Mins = $row_rsRadius['sfCycle_Mins'];
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("Update accountinfo Set sfCycle_Mins = sfCycle_Mins - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Mins'],$row_cyclemet['acci_RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Mins = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
											}
											//1.9
											if  ($row_cyclemet['PeriodFee'] == 0)
											{
												$cCharge = $cCharge + $row_plantype['PeriodFee'];
											} else {
												$cCharge = $cCharge + $row_cyclemet['PeriodFee'];
											}
				 
											 
											mysql_select_db($database_projectalpha, $projectalpha);
											$query_sqlv = sprintf("select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = '%s'",$row_cyclemet['acci_RecID']);
											$rsprepayment = mysql_query($query_sqlv, $projectalpha) or die(mysql_error());
											$totalRows_rsprepayment = mysql_num_rows($rsprepayment);
											
						
											$cPrePaid = 0;
											$cTotalDue = 0;
											$cAmountPaid = 0;
											$AddReceipt = false;
											$cTotalDue = $cCharge + $cCharge * vidtax($row_cyclemet['VirtualID']);
											//1.9
											if ($totalRows_rsprepayment > 0)
											{
												//1.10
												while ($row_rsprepayment = mysql_fetch_assoc($rsprepayment) || $cTotalDue == 0) 
													{
													//1.11
													 if (($row_rsprepayment['AmountPaid'] - $row_rsprepayment['AmountUsed']) > 0)
														{
														$cPrePaid = $row_rsprepayment['AmountPaid'] - $row_rsprepayment['AmountUsed'];
														//1.12
														if ($cPrePaid < $cTotalDue) {
															$cTotalDue = $cTotalDue - $cPrePaid;
															$cAmountPaid = $cAmountPaid + $cPrePaid;
															mysql_select_db($database_projectalpha, $projectalpha);
															$query_sqlg = sprintf("UPDATE invoicein Set AmountUsed = AmountPaid where RecID = '%s'",$row_rsprepayment['RecID']);
															$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
															$AddReceipt=true;
														} else {
															//1.13
															if ($cPrePaid > $cTotalDue)
															{
													  
															mysql_select_db($database_projectalpha, $projectalpha);
															$query_sqlg = sprintf("UPDATE invoicein Set AmountUsed = AmountUsed + %d where RecID = '%s'",$cTotalDue,$row_rsprepayment['RecID']);
															$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
						
															$cAmountPaid = $cAmountPaid + $cTotalDue;
															$cTotalDue = 0;
															$AddReceipt=true;
															}
															//1.13
														}
														//1.12
													}
													//1.11
												}
												//1.10
											}
											//1.9
											if ($AddReciept=true) {
												$invout_PaidWhen = $row_datenav['DateNava'];
											} else {
												$invout_PaidWhen = 'NULL';
											}
											//1.9
											$invout_AmountDue = $cCharge;
											$invout_GSTCharged = $cCharge  * vidtax($row_cyclemet['VirtualID']);
											$invout_TotalDue = $cTotalDue;
											$invout_AmountRefunded = 0;
											$invout_GSTRefunded = 0;
											$invout_AmountPaid = $cAmountPaid;
											$invout_acci_RecID = $row_cyclemet['acci_RecID'];
											$invout_SysopID = $row_cyclemet['SysopID'];
											$invout_PlanServiceID = $row_cyclemet['RecID'];
						
											$sqlb = sprintf("select RecID, VirtualID, payIntervalType, payInterval from accountinfo Where RecID = '%d'" , $cyclemet['acci_RecID']);
											
											mysql_select_db($database_projectalpha, $projectalpha);
											//$query_sqlf = sprintf("select RecID, sfCycle_Upload, sfCycle_Download, sfCycle_Mins from accountinfo Where RecID = '%s'",$row_cyclemet['acci_RecID']);
											$rsPaymentInt = mysql_query($sqlb, $projectalpha) or die(mysql_error());
											$totalRows_rsPaymentInt = mysql_num_rows($rsPaymentInt);
											$row_rsPaymentInt = mysql_fetch_assoc($rsPaymentInt);
											//1.9
											if ($totalRows_rsPaymentInt==0)	{
												$phpdateadd = "14 days";
											} else {
												//1.10
												if ($row_rsPaymentInt['payIntervalType']=='d') {
													$phpdateadd = sprintf("%d day",$row_rsPaymentInt['payInterval']);
													if ($row_rsPaymentInt['payInterval']>1) {
														$phpdateadd .= "s";
													}
												} else {
													//1.11
													if ($row_rsPaymentInt['payIntervalType']=='h') {
														$phpdateadd = sprintf("%d hour",$row_rsPaymentInt['payInterval']);
														if ($row_rsPaymentInt['payInterval']>1) {
															$phpdateadd .= "s";
														}
													} else {
														//1.12
														if ($row_rsPaymentInt['payIntervalType']=='m') {
															$phpdateadd = sprintf("%d month",$row_rsPaymentInt['payInterval']);
															if ($row_rsPaymentInt['payInterval']>1) {
																$phpdateadd .= "s";
															}
						
														} else {
															//1.13
															if ($row_rsPaymentInt['payIntervalType']=='q') {
																$phpdateadd = sprintf("%d month",$row_rsPaymentInt['payInterval']*4);
																if ($row_rsPaymentInt['payInterval']>=1) {
																	$phpdateadd .= "s";
																}
						
															} else {
																//1.14
																if ($row_rsPaymentInt['payIntervalType']=='ww' || $row_rsPaymentInt['payIntervalType']=="w") {
																	$phpdateadd = sprintf("%d day",$row_rsPaymentInt['payInterval']*7);
																	if ($row_rsPaymentInt['payInterval']>=1) {
																		$phpdateadd .= "s";
																	}
																} else {
																	//1.15
																	if ($row_rsPaymentInt['payIntervalType']=='y' || $row_rsPaymentInt['payIntervalType']=="yyyy") {
																		$phpdateadd = sprintf("%d year",$row_rsPaymentInt['payInterval']);
																		if ($row_rsPaymentInt['payInterval']>1) {
																			$phpdateadd .= "s";
																		}
																	} else {
																		$phpdateadd = "14 days";
																	}
																	//1.15
																}
																//1.14
															}
															//1.13
														}
														//1.12
													}
													//1.11
												}
												//1.10
											}
											//1.9
												
											//1.9
											if ($totalRows_rsPaymentInt > 0) {
												//$invout.PaymentDue = DateAdd(Iif (IsNull(rsload2!PayIntervalType), "d", rsload2!PayIntervalType), Iif (IsNull(rsload2!PayInterval), 14, rsload2!PayInterval), sysnow)
												$invout_VirtualID = $row_rsPaymentInt['VirtualID'];
											} else {
												//                       invout.PaymentDue = DateAdd("d", 14, sysnow)
												$invout_VirtualID = $_SESSION['VirtualID'];
											}
											
											//1.9
											if ($bNewGroup == true) {
												
												$NumInvCreated = $NumInvCreated + 1;
													
												$sqla = sprintf("insert into invoiceout (Description, PaidWhen) Values('%s','%s')",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												$sqlb = sprintf("select RecID from invoiceout where Description = '%s' and PaidWhen = '%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlb = mysql_query($sqlb, $projectalpha) or die(mysql_error());
												$totalRows_rssqlb = mysql_num_rows($rssqlb);
												$row_rssqlb = mysql_fetch_assoc($rssqlb);
										
												$invout_RecID = $row_rssqlb['RecID'];
												$invout_SubRecID = $row_rssqlb['RecID'];
												
												$insertsql = "update invoiceout set Description = '%s', PaidWhen = '%s', AmountDue = '%d', GSTCharged = '%d', ";
												$insertsql .= "TotalDue = '%d', AmountRefunded = '%d', GSTRefunded = '%d', AmountPaid = '%d', acci_RecID = '%d', SysopID = '%d', PlanServiceID = '%d', SubRecID = '%d', StartCycle = '%s', EndCycle = '%s', PaymentDue = dateadd(now(), '%s'), VirtualID = '%d', StatementID = '%d', sfCycle_Download = '%d', sfCycle_Upload = '%d', ";
												$insertsql .= "sfCycle_Mins = '%d', StatementID = '%d' where RecID  = '%d'";
												$sqlc = sprintf($insertsql, $invout_Description, $invout_PaidWhen, $invout_AmountDue, $invout_GSTCharged, $invout_TotalDue, $invout_AmountRefunded, $invout_GSTRefunded, $invout_AmountPaid, $invout_acci_RecID, $invout_SysopID, $invout_PlanServiceID, $invout_SubRecID, $invout_StartCycle, $invout_EndCycle, $phpdateadd, $invout_VirtualID, $invout_StatementID, $invout_sfCycle_Download, $invout_sfCycle_Upload, $invout_sfCycle_Mins, $invout_StatementID, $invout_RecID);
									 
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlc = mysql_query($sqlc, $projectalpha) or die(mysql_error());
											 
												$sqld = sprintf("insert into statementitems (Description) VALUES (%s%s)",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												$sqle = sprintf("select RecID from statementitems where Description ='%s%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqld = mysql_query($sqld, $projectalpha) or die(mysql_error());
												
												//$StatementID = MySQL.GetTMPRecID("statementitems", oConn, "RecID", False)
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqle = mysql_query($sqle, $projectalpha) or die(mysql_error());
												$totalRows_rssqle = mysql_num_rows($rssqle);
												$row_rssqle = mysql_fetch_assoc($rssqle);
												
												$StatementID = $row_rssqle['RecID'];
												
												$sqlf = sprintf("update statementitems set InvRecID = '%d', Items = '1', Description = '%s', TotalDue = '%d' where RecID = '%d'",$invout_SubRecID,$invout_Description, $cTotalDue, $StatementID);
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlf = mysql_query($sqlf, $projectalpha) or die(mysql_error());
												
												$sqlg = sprintf("update invoiceout Set StatementID = '%d' where RecID = '$d'",$StatementID,$invout_SubRecID);
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlg = mysql_query($sqlg, $projectalpha) or die(mysql_error());
												
												$sqlh = sprintf("update online_maintenancelocker set Active = 2 where SvrRecID = '%d' and SysopIDRunningMaintenance = '%d' and VirtualID = '%d and DateNava = '%s', DateNavb = '%s''",$row_cyclemet['RecID'], $_SESSION['SysopID'], $row_cyclemet['VirtualID'], $row_datenav['DateNava'], $row_datenav['DateNavb']); 
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlh = mysql_query($sqlh, $projectalpha) or die(mysql_error());
				
												$bNewGroup = false;
												
											} else {
											
												$sqla = sprintf("insert into invoiceout (Description, PaidWhen) Values('%s','%s')",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												$sqlb = sprintf("select RecID from invoiceout where Description = '%s' and PaidWhen = '%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlb = mysql_query($sqlb, $projectalpha) or die(mysql_error());
												$totalRows_rssqlb = mysql_num_rows($rssqlb);
												$row_rssqlb = mysql_fetch_assoc($rssqlb);
										
												$invout_RecID = $row_rssqlb['RecID'];
												$invout_StatementID = $StatementID;
												
												$insertsql = "update invoiceout set Description = '%s', PaidWhen = '%s', AmountDue = '%d', GSTCharged = '%d', ";
												$insertsql .= "TotalDue = '%d', AmountRefunded = '%d', GSTRefunded = '%d', AmountPaid = '%d', acci_RecID = '%d', SysopID = '%d', PlanServiceID = '%d', SubRecID = '%d', StartCycle = '%s', EndCycle = '%s', PaymentDue = dateadd(now(), '%s'), VirtualID = '%d', StatementID = '%d', sfCycle_Download = '%d', sfCycle_Upload = '%d', ";
												$insertsql .= "sfCycle_Mins = '%d', StatementID = '%d' where RecID  = '%d'";
												$sqlc = sprintf($insertsql, $invout_Description, $invout_PaidWhen, $invout_AmountDue, $invout_GSTCharged, $invout_TotalDue, $invout_AmountRefunded, $invout_GSTRefunded, $invout_AmountPaid, $invout_acci_RecID, $invout_SysopID, $invout_PlanServiceID, $invout_SubRecID, $invout_StartCycle, $invout_EndCycle, $phpdateadd, $invout_VirtualID, $invout_StatementID, $invout_sfCycle_Download, $invout_sfCycle_Upload, $invout_sfCycle_Mins, $invout_StatementID, $invout_RecID);
									 
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlc = mysql_query($sqlc, $projectalpha) or die(mysql_error());
											
												$sqld = sprintf("Update statementitems set Items=Items+1,TotalDue = TotalDue + '%d' where RecID = '%d'",$cTotalDue, $invout_StatementID);
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqld = mysql_query($sqld, $projectalpha) or die(mysql_error());
											
												$sqlh = sprintf("update online_maintenancelocker set Active = 2 where SvrRecID = '%d' and SysopIDRunningMaintenance = '%d' and VirtualID = '%d and DateNava = '%s', DateNavb = '%s''",$row_cyclemet['RecID'], $_SESSION['SysopID'], $row_cyclemet['VirtualID'], $row_datenav['DateNava'], $row_datenav['DateNavb']);
												mysql_select_db($database_projectalpha, $projectalpha);
												$rssqlh = mysql_query($sqlh, $projectalpha) or die(mysql_error());
												
											}
											//1.9
											if ( $AddReceipt = true ) {
												AddReceiptItem( $invout_acci_RecID, $invout_RecID,  $_SESSION['SysopID'], $row_cyclemet['RecID'], $invout_AmountPaid, $invout_AmountRefunded,  $invout_GSTCharged, $invout_GSTRefunded, $row_datenav['DateNava'], 'Account Payment from the Vault','Prepayment/The Vault','', 0, 0, $StatementID);
											}
											//1.9
										}
										//1.8
										$cChargeTotal = $cChargeTotal + $cCharge;
									}  
									// 1.7
							} else {
							//1.6
								
								$cCharge = 0;
				
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_sqlb = sprintf("select * from plantypes where RecID ='%s'",$row_cyclemet['ptRecID']);
								$plantype = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
								$totalRows_plantype = mysql_num_rows($plantype);
								$row_plantype = mysql_fetch_assoc($plantype);
								//1.7
								if ($totalRows_plantype > 0)
									{
								 
										mysql_select_db($database_projectalpha, $projectalpha);
										$query_sqlc = sprintf("select AccountActive from acci_dslconnections where acci_RecID = '%s'",$row_cyclemet['acci_RecID']);
										$accactive = mysql_query($query_sqlc, $projectalpha) or die(mysql_error());
										$totalRows_accactive = mysql_num_rows($accactive);
										$row_accactive = mysql_fetch_assoc($accactive);
										//1.8
										if ($totalRows_accactive > 0) 
										{
											//1.9
											if ($row_plantype['BillOnce'] ==1 || $row_cyclemet['Activation'] == $row_cyclemet['NextCycle']) 
											{
				
												$invout_Description = $plantype['Description'];
					
												//1.10
												if ($cyclemet['Activation'] == $row_cyclemet['NextCycle']) 
													{ 
					
														$invout_Description .= sprintf(" - Account Setup Fee [%.2f",$row_cyclemet['JoiningFee']);
														$cCharge = $cCharge + $row_cyclemet['JoiningFee'];
														mysql_select_db($database_projectalpha, $projectalpha);
														$query_rsVISPb = sprintf("Update acci_services set Checked='-1' where RecID = ",$row_cyclemet['RecID']);
														$datenav = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
					
													}
														
					
												//1.10
												if ($row_cyclemet['MBPerPeriod'] <> -1) 
													{
													if ( $row_rsRadius['sfCycle_Download'] / $lBytesPerMB > $row_plantype['MBPerPeriod']) {
														$invout_Description .= sprintf(" %.2f MB's Over", ($row_rsRadius['sfCycle_Download'] / $lBytesPerMB) -  $row_plantype['MBPerPeriod']);
														$cCharge = $cCharge + (($row_rsRadius['sfCycle_Download'] / $lBytesPerMB -  $row_plantype['MBPerPeriod']) /  $row_plantype['MBBlockSize']) * $row_cyclemet['PerMB'];
													}
												
												$invout_sfCycle_Download = $row_rsRadius['sfCycle_Download'];
												$invout_sfCycle_Upload = $row_rsRadius['sfCycle_Upload'];
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlf = sprintf("select RecID, sfCycle_Upload, sfCycle_Download, sfCycle_Mins from accountinfo Where RecID = '%s'",$row_cyclemet['acci_RecID']);
												$rsAccInfo = mysql_query($query_sqlf, $projectalpha) or die(mysql_error());
												$totalRows_rsAccInfo = mysql_num_rows($accactive);
												$row_rsAccInfo = mysql_fetch_assoc($accactive);
												//1.10
												if ($totalRows_rsAccInfo > 0)
												{
							
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Insert Into history_acci_datausage (acci_RecID, Uploaded, Downloaded, NumMins) VALUES('%s','%d','%d','%d')",$row_cyclemet['acci_RecID'],$row_rsAccInfo['sfCycle_Upload'],$row_rsAccInfo['sfCycle_Download'],$row_rsAccInfo['sfCycle_Mins']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Upload = sfCycle_Upload - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Upload'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Download = sfCycle_Download - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Download'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
										
													mysql_select_db($database_projectalpha, $projectalpha);
													$query_sqlg = sprintf("Update accountinfo Set sfCycle_Upload = sfCycle_Upload - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Mins'],$row_cyclemet['acci_RecID']);
													$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												   
												}
												//1.10
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Download = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Upload = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												//1.10
												if ($row_plantype['HoursPerPeriod'] <> -1)
												{
													//1.11
													if ($row_rsRadius['sfCycle_Mins'] / 60 > $row_plantype['HoursPerPeriod'])
													{
														$invout_Description = sprintf("%s %d Min's Over",$invout_Description, ($row_rsRadius['sfCycle_Mins'] - $row_plantype['HoursPerPeriod'] * 60));
														$cCharge = $cCharge + ($row_rsRadius['sfCycle_Mins'] / 60) * $row_cyclemet['PerHour'];
													} //1.11
												} //1.10
												
												$invout_sfCycle_Mins = $row_rsRadius['sfCycle_Mins'];
									
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("Update accountinfo Set sfCycle_Mins = sfCycle_Mins - '%d' where RecID = '%s'",$row_rsRadius['sfCycle_Mins'],$row_cyclemet['acci_RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlg = sprintf("UPDATE radiusaccounts Set sfCycle_Mins = '%d' where RecID = '%s'",0,$row_rsRadius['RecID']);
												$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
												//1.10
												if  ($row_cyclemet['PeriodFee'] == 0)
												{
													$cCharge = $cCharge + $row_plantype['PeriodFee'];
												} else {
													$cCharge = $cCharge + $row_cyclemet['PeriodFee'];
												}
				 
												mysql_select_db($database_projectalpha, $projectalpha);
												$query_sqlv = sprintf("select * from invoicein Where AmountPaid > AmountUsed AND AccI_RecID = '%s'",$row_cyclemet['acci_RecID']);
												$rsprepayment = mysql_query($query_sqlv, $projectalpha) or die(mysql_error());
												$totalRows_rsprepayment = mysql_num_rows($rsprepayment);
												
							
												$cPrePaid = 0;
												$cTotalDue = 0;
												$cAmountPaid = 0;
												$AddReceipt = false;
												$cTotalDue = $cCharge + $cCharge * vidtax($row_cyclemet['VirtualID']);
												//1.10
												if ($totalRows_rsprepayment > 0)
												{
													//1.11
													while ($row_rsprepayment = mysql_fetch_assoc($rsprepayment) || $cTotalDue == 0) 
														{
														 //1.12	
														 if (($row_rsprepayment['AmountPaid'] - $row_rsprepayment['AmountUsed']) > 0)
															{
															$cPrePaid = $row_rsprepayment['AmountPaid'] - $row_rsprepayment['AmountUsed'];
															//1.13
															if ($cPrePaid < $cTotalDue) {
																$cTotalDue = $cTotalDue - $cPrePaid;
																$cAmountPaid = $cAmountPaid + $cPrePaid;
																mysql_select_db($database_projectalpha, $projectalpha);
																$query_sqlg = sprintf("UPDATE invoicein Set AmountUsed = AmountPaid where RecID = '%s'",$row_rsprepayment['RecID']);
																$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
																$AddReceipt=true;
															} else {
															    //1.14
																if ($cPrePaid > $cTotalDue)
																{
														  
																mysql_select_db($database_projectalpha, $projectalpha);
																$query_sqlg = sprintf("UPDATE invoicein Set AmountUsed = AmountUsed + %d where RecID = '%s'",$cTotalDue,$row_rsprepayment['RecID']);
																$execute = mysql_query($query_sqlg, $projectalpha) or die(mysql_error());
							
																$cAmountPaid = $cAmountPaid + $cTotalDue;
																$cTotalDue = 0;
																$AddReceipt=true;
																}
																//1.14
															}
															//1.13
														}
														//1.12
													}
													//1.11
												}
												//1.10
												if ($AddReciept=true) {
													$invout_PaidWhen = $row_datenav['DateNava'];
												} else {
													$invout_PaidWhen = 'NULL';
												}
									
												$invout_AmountDue = $cCharge;
												$invout_GSTCharged = $cCharge  * vidtax($row_cyclemet['VirtualID']);
												$invout_TotalDue = $cTotalDue;
												$invout_AmountRefunded = 0;
												$invout_GSTRefunded = 0;
												$invout_AmountPaid = $cAmountPaid;
												$invout_acci_RecID = $row_cyclemet['acci_RecID'];
												$invout_SysopID = $row_cyclemet['SysopID'];
												$invout_PlanServiceID = $row_cyclemet['RecID'];
				
												$sqlb = sprintf("select RecID, VirtualID, payIntervalType, payInterval from accountinfo Where RecID = '%d'" , $cyclemet['acci_RecID']);
												
												mysql_select_db($database_projectalpha, $projectalpha);
												//$query_sqlf = sprintf("select RecID, sfCycle_Upload, sfCycle_Download, sfCycle_Mins from accountinfo Where RecID = '%s'",$row_cyclemet['acci_RecID']);
												$rsPaymentInt = mysql_query($sqlb, $projectalpha) or die(mysql_error());
												$totalRows_rsPaymentInt = mysql_num_rows($rsPaymentInt);
												$row_rsPaymentInt = mysql_fetch_assoc($rsPaymentInt);
												//1.10
												if ($totalRows_rsPaymentInt==0)	{
													$phpdateadd = "14 days";
												} else {
													//1.11
													if ($row_rsPaymentInt['payIntervalType']=='d') {
														$phpdateadd = sprintf("%d day",$row_rsPaymentInt['payInterval']);
														//1.12
														if ($row_rsPaymentInt['payInterval']>1) {
															$phpdateadd .= "s";
														}
													} else {
													    //1.12
														if ($row_rsPaymentInt['payIntervalType']=='h') {
															$phpdateadd = sprintf("%d hour",$row_rsPaymentInt['payInterval']);
															//1.13
															if ($row_rsPaymentInt['payInterval']>1) {
																$phpdateadd .= "s";
															}
														} else {
															//1.13
															if ($row_rsPaymentInt['payIntervalType']=='m') {
																$phpdateadd = sprintf("%d month",$row_rsPaymentInt['payInterval']);
																//1.14
																if ($row_rsPaymentInt['payInterval']>1) {
																	$phpdateadd .= "s";
																}
							
															} else {
																//1.14
																if ($row_rsPaymentInt['payIntervalType']=='q') {
																	$phpdateadd = sprintf("%d month",$row_rsPaymentInt['payInterval']*4);
																	//1.15
																	if ($row_rsPaymentInt['payInterval']>=1) {
																		$phpdateadd .= "s";
																	}
							
																} else {
																	//1.15
																	if ($row_rsPaymentInt['payIntervalType']=='ww' || $row_rsPaymentInt['payIntervalType']=="w") {
																		$phpdateadd = sprintf("%d day",$row_rsPaymentInt['payInterval']*7);
																		//1.16
																		if ($row_rsPaymentInt['payInterval']>=1) {
																			$phpdateadd .= "s";
																		}
																	} else {
																		//1.16
																		if ($row_rsPaymentInt['payIntervalType']=='y' || $row_rsPaymentInt['payIntervalType']=="yyyy") {
																			$phpdateadd = sprintf("%d year",$row_rsPaymentInt['payInterval']);
																			//1.17
																			if ($row_rsPaymentInt['payInterval']>1) {
																				$phpdateadd .= "s";
																			}
																		} else {
																			//1.17
																			$phpdateadd = "14 days";
																		}
																		//1.16
																	}
																	//1.15
																}
																//1.14
															}
															//1.13
														}
														//1.12

													}
													//1.11
														
												}
												//1.10
									
									
												if ($totalRows_rsPaymentInt > 0) {
													//$invout.PaymentDue = DateAdd(Iif (IsNull(rsload2!PayIntervalType), "d", rsload2!PayIntervalType), Iif (IsNull(rsload2!PayInterval), 14, rsload2!PayInterval), sysnow)
													$invout_VirtualID = $row_rsPaymentInt['VirtualID'];
												} else {
													//                       invout.PaymentDue = DateAdd("d", 14, sysnow)
													$invout_VirtualID = $_SESSION['VirtualID'];
												}
												
									
												//1.10
												if ($bNewGroup == true) {
													$NumInvCreated = $NumInvCreated + 1;
														
														
													$sqla = sprintf("insert into invoiceout (Description, PaidWhen) Values('%s','%s')",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													$sqlb = sprintf("select RecID from invoiceout where Description = '%s' and PaidWhen = '%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlb = mysql_query($sqlb, $projectalpha) or die(mysql_error());
													$totalRows_rssqlb = mysql_num_rows($rssqlb);
													$row_rssqlb = mysql_fetch_assoc($rssqlb);
											
													$invout_RecID = $row_rssqlb['RecID'];
													$invout_SubRecID = $row_rssqlb['RecID'];
													
													$insertsql = "update invoiceout set Description = '%s', PaidWhen = '%s', AmountDue = '%d', GSTCharged = '%d', ";
													$insertsql .= "TotalDue = '%d', AmountRefunded = '%d', GSTRefunded = '%d', AmountPaid = '%d', acci_RecID = '%d', SysopID = '%d', PlanServiceID = '%d', SubRecID = '%d', StartCycle = '%s', EndCycle = '%s', PaymentDue = dateadd(now(), '%s'), VirtualID = '%d', StatementID = '%d', sfCycle_Download = '%d', sfCycle_Upload = '%d', ";
													$insertsql .= "sfCycle_Mins = '%d', StatementID = '%d' where RecID  = '%d'";
													$sqlc = sprintf($insertsql, $invout_Description, $invout_PaidWhen, $invout_AmountDue, $invout_GSTCharged, $invout_TotalDue, $invout_AmountRefunded, $invout_GSTRefunded, $invout_AmountPaid, $invout_acci_RecID, $invout_SysopID, $invout_PlanServiceID, $invout_SubRecID, $invout_StartCycle, $invout_EndCycle, $phpdateadd, $invout_VirtualID, $invout_StatementID, $invout_sfCycle_Download, $invout_sfCycle_Upload, $invout_sfCycle_Mins, $invout_StatementID, $invout_RecID);
										 
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlc = mysql_query($sqlc, $projectalpha) or die(mysql_error());
												   
													//$StatementID = MySQL.GetTMPRecID("statementitems", oConn, "RecID", False)
													$sqld = sprintf("insert into statementitems (Description) VALUES (%s%s)",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													$sqle = sprintf("select RecID from statementitems where Description ='%s%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqld = mysql_query($sqld, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqle = mysql_query($sqle, $projectalpha) or die(mysql_error());
													$totalRows_rssqle = mysql_num_rows($rssqle);
													$row_rssqle = mysql_fetch_assoc($rssqle);
													
													$StatementID = $row_rssqle['RecID'];
													
													$sqlf = sprintf("update statementitems set InvRecID = '%d', Items = '1', Description = '%s', TotalDue = '%d' where RecID = '%d'",$invout_SubRecID,$invout_Description, $cTotalDue, $StatementID);
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlf = mysql_query($sqlf, $projectalpha) or die(mysql_error());
													
													$sqlg = sprintf("update invoiceout Set StatementID = '%d' where RecID = '$d'",$StatementID,$invout_SubRecID);
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlg = mysql_query($sqlg, $projectalpha) or die(mysql_error());
													$bNewGroup = false;
													
													$sqlh = sprintf("update online_maintenancelocker set Active = 2 where SvrRecID = '%d' and SysopIDRunningMaintenance = '%d' and VirtualID = '%d and DateNava = '%s', DateNavb = '%s''",$row_cyclemet['RecID'], $_SESSION['SysopID'], $row_cyclemet['VirtualID'], $row_datenav['DateNava'], $row_datenav['DateNavb']); 
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlh = mysql_query($sqlh, $projectalpha) or die(mysql_error());
													
												} else {
												
													$sqla = sprintf("insert into invoiceout (Description, PaidWhen) Values('%s','%s')",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													$sqlb = sprintf("select RecID from invoiceout where Description = '%s' and PaidWhen = '%s'",$row_datenav['DateNava'],$row_datenav['DateNavb']);
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlb = mysql_query($sqlb, $projectalpha) or die(mysql_error());
													$totalRows_rssqlb = mysql_num_rows($rssqlb);
													$row_rssqlb = mysql_fetch_assoc($rssqlb);
											
													$invout_RecID = $row_rssqlb['RecID'];
													$invout_StatementID = $StatementID;
													
													$insertsql = "update invoiceout set Description = '%s', PaidWhen = '%s', AmountDue = '%d', GSTCharged = '%d', ";
													$insertsql .= "TotalDue = '%d', AmountRefunded = '%d', GSTRefunded = '%d', AmountPaid = '%d', acci_RecID = '%d', SysopID = '%d', PlanServiceID = '%d', SubRecID = '%d', StartCycle = '%s', EndCycle = '%s', PaymentDue = dateadd(now(), '%s'), VirtualID = '%d', StatementID = '%d', sfCycle_Download = '%d', sfCycle_Upload = '%d', ";
													$insertsql .= "sfCycle_Mins = '%d', StatementID = '%d' where RecID  = '%d'";
													$sqlc = sprintf($insertsql, $invout_Description, $invout_PaidWhen, $invout_AmountDue, $invout_GSTCharged, $invout_TotalDue, $invout_AmountRefunded, $invout_GSTRefunded, $invout_AmountPaid, $invout_acci_RecID, $invout_SysopID, $invout_PlanServiceID, $invout_SubRecID, $invout_StartCycle, $invout_EndCycle, $phpdateadd, $invout_VirtualID, $invout_StatementID, $invout_sfCycle_Download, $invout_sfCycle_Upload, $invout_sfCycle_Mins, $invout_StatementID, $invout_RecID);
										 
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlc = mysql_query($sqlc, $projectalpha) or die(mysql_error());
												
													$sqld = sprintf("Update statementitems set Items=Items+1,TotalDue = TotalDue + '%d' where RecID = '%d'",$cTotalDue, $invout_StatementID);
													
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqld = mysql_query($sqld, $projectalpha) or die(mysql_error());
												
													$sqlh = sprintf("update online_maintenancelocker set Active = 2 where SvrRecID = '%d' and SysopIDRunningMaintenance = '%d' and VirtualID = '%d and DateNava = '%s', DateNavb = '%s''",$row_cyclemet['RecID'], $_SESSION['SysopID'], $row_cyclemet['VirtualID'], $row_datenav['DateNava'], $row_datenav['DateNavb']); 
													mysql_select_db($database_projectalpha, $projectalpha);
													$rssqlh = mysql_query($sqlh, $projectalpha) or die(mysql_error());
													
												}
												//1.10
												
												if ( $AddReceipt = true ) {
													AddReceiptItem( $invout_acci_RecID, $invout_RecID,  $_SESSION['SysopID'], $row_cyclemet['RecID'], $invout_AmountPaid, $invout_AmountRefunded,  $invout_GSTCharged, $invout_GSTRefunded, $row_datenav['DateNava'], 'Account Payment from the Vault','Prepayment/The Vault','', 0, 0, $StatementID);
												}
												
												$cTotalCharged = $cTotalCharged + $cCharge;
											}
											//1.9
											
										}
										//1.8
									}
									//1.7
								}
								//1.6
		
							mysql_select_db($database_projectalpha, $projectalpha);
							$query_sqll= sprintf("select chgIntervalType, chgInterval from plantypes where RecID ='%s'",$row_cyclemet['ptRecID']);
							$nextcycle = mysql_query($query_rsVISPl, $projectalpha) or die(mysql_error());
							$totalRows_nextcycle = mysql_num_rows($nextcycle);
							$row_nextcycle = mysql_fetch_assoc($nextcycle);
						
							//1.5
							if ($totalRows_nextcycle==0)	{
								$phpdateadd = "1 month";
							} else {
								//1.6
								if ($row_nextcycle['chgIntervalType']=='d') {
									$phpdateadd = sprintf("%d day",$row_nextcycle['chgInterval']);
									if ($row_nextcycle['chgInterval']>1) {
										$phpdateadd .= "s";
									}
								} else {
									//1.7
									if ($row_nextcycle['chgIntervalType']=='h') {
										$phpdateadd = sprintf("%d hour",$row_nextcycle['chgInterval']);
										if ($row_nextcycle['chgInterval']>1) {
											$phpdateadd .= "s";
										}
									} else {
										//1.8
										if ($row_nextcycle['chgIntervalType']=='m') {
											$phpdateadd = sprintf("%d month",$row_nextcycle['chgInterval']);
											if ($row_nextcycle['chgInterval']>1) {
												$phpdateadd .= "s";
											}
				
										} else {
											//1.9
											if ($row_nextcycle['chgIntervalType']=='q') {
												$phpdateadd = sprintf("%d month",$row_nextcycle['chgInterval']*4);
												if ($row_nextcycle['chgInterval']>=1) {
													$phpdateadd .= "s";
												}
				
											} else {
												//1.10
												if ($row_nextcycle['chgIntervalType']=='ww' || $row_nextcycle['chgIntervalType']=="w") {
													$phpdateadd = sprintf("%d day",$row_nextcycle['chgInterval']*7);
													if ($row_nextcycle['chgInterval']>=1) {
														$phpdateadd .= "s";
													}
												} else {
													//1.11
													if ($row_nextcycle['chgIntervalType']=='y' || $row_nextcycle['chgIntervalType']=="yyyy") {
														$phpdateadd = sprintf("%d year",$row_nextcycle['chgInterval']);
														if ($row_nextcycle['chgInterval']>1) {
															$phpdateadd .= "s";
														}
													} else {
														$phpdateadd = "1 month";
													}
													//1.11
												}
												//1.10
											}
											//1.9
										}
										//1.8
									}
									//1.7
								}
								//1.6
							}
							//1.5
								
							
							if ($totalRows_offset > 0) {
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_rsVISPb = sprintf("UPDATE acci_services Set NextCycle = dateadd(NextCycle, '\"%d seconds\") where RecID = '%d'",$row_offset['IntegerSet'],$row_cyclemet['RecID']);
								$datenavz = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
								//$row_datenav = mysql_fetch_assoc($datenav);
							} else {
								mysql_select_db($database_projectalpha, $projectalpha);
								$query_rsVISPb = sprintf("UPDATE acci_services Set NextCycle = dateadd(NextCycle, '\"%d seconds\") where RecID = '%d'",30,$row_cyclemet['RecID']);
								$datenavz = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
								//$row_datenav = mysql_fetch_assoc($datenav);
							}			
							
							$query_nextcycle2 = sprintf("select dateadd(NextCycle, '%s') as nextcycle from acci_services where RecID = ",$row_cyclemet['RecID'],$phpdateadd);
							mysql_select_db($database_projectalpha, $projectalpha);
							$nextcycle2 = mysql_query($query_nextcycle2, $projectalpha) or die(mysql_error());
							$row_nextcycle2 = mysql_fetch_assoc($nextcycle2);
				
							$newnextcycle = date("M d Y",$row_nextcycle2['nextcycle']);
							$newnextcycle .= " ";
							if (phpversion() >= 5) {
								$newnextcycle .= date("H:i:s", date_sunrise($row_nextcycle2['nextcycle'], SUNFUNCS_RET_STRING, $row_LongandLatit['Latitude'], $row_LongandLatit['Longitude'], 0));
							} else {
								$newnextcycle .= date("H:i:s", $row_nextcycle2['nextcycle']);
							}
							
							mysql_select_db($database_projectalpha, $projectalpha);
							$query_UpdateCycle = sprintf("update acci_services set NextCycle = '%s', PreviousCycle=NextCycle, PreviousCycleA=PreviousCycle, PreviousCycleB=PreviousCycleA, PreviousCycleC=PreviousCycleB, PreviousCycleD=PreviousCycleC, PreviousCycleE=PreviousCycleD, PreviousCycleF=PreviousCycleE where RecID = '%d'",$newnextcycle, $row_cyclemet['RecID']);
							$update = mysql_query($query_UpdateCycle, $projectalpha) or die(mysql_error());
				
							mysql_select_db($database_projectalpha, $projectalpha);
							$query_UpdateCycle = sprintf("update online_maintenancelocker set Active = 2, StatementID = '%d' where RecID = '%d'", $StatementID, $row_cyclemet['RecID']);
							$update = mysql_query($query_UpdateCycle, $projectalpha) or die(mysql_error());			
							
							
				
						}
						//1.4
					
					}
					//1.3
				
					mysql_select_db($database_projectalpha, $projectalpha);
					$query_rsVISPb = sprintf("select IntegerSet from  where ConfigKey = 'DOINVOICEOFFSET' and VirtualID ='%d'",$_SESSION['VirtualID']);
					$offset2 = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
					$totalRows_offset2 = mysql_num_rows($offset2);
					//1.3
					if ($totalRows_offset2==0) {
						mysql_select_db($database_projectalpha, $projectalpha);
						$query_rsVISPb = sprintf("select NOW() as DateNava, dateadd(now(), \"%d seconds\") as DateNavb",30);
						$datenav2 = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
						$row_datenav2 = mysql_fetch_assoc($datenav2);
						
						$sqla = sprintf("insert into online_settingsandtimers (IntegerSet, ConfigKey, VirtualID, SysopID) Values ('-%d','DOINVOICEOFFSET','%d','%d')", dateDiff('s',$row_datenav['DateNava'],$row_datenav2['DateNava']), $_SESSION['VirtualID'], $_SESSION['SysopID']);
						mysql_select_db($database_projectalpha, $projectalpha);
						$execute = mysql_query($sqla, $projectalpha) or die(mysql_error());	
					} else {
						mysql_select_db($database_projectalpha, $projectalpha);
						$query_rsVISPb = sprintf("select NOW() as DateNava, dateadd(now(), \"%d seconds\") as DateNavb",$row_offset['IntegerSet']);
						$datenav2 = mysql_query($query_rsVISPb, $projectalpha) or die(mysql_error());
						$row_datenav2 = mysql_fetch_assoc($datenav2);
						
						$sqla = sprintf("Update online_settingsandtimers set IntegerSet = '-%d' where ConfigKey = 'DOINVOICEOFFSET' and VirtualID = '%d'", dateDiff('s',$row_datenav['DateNava'],$row_datenav2['DateNava']), $_SESSION['VirtualID']);
						mysql_select_db($database_projectalpha, $projectalpha);
						$sqla = mysql_query($sqla, $projectalpha) or die(mysql_error());
					}	
					//1.3
				}
				//1.2
				$DisplayProcessResults=true;
			}
			//1.1
	}
	}

?>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Generated <?php echo sprintf(" %d Invoices - A total transaction value of [%.2f",$cTotalCharged); ?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: small;
	color: #FF9933;
}
body {
	background-color: #666633;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 10px;
	background-image: url(images/bricks%20by%20moonlight.jpg);
}
.style1 {
	font-size: x-large;
	color: #000066;
}
.style2 {font-weight: bold}
-->
</style></head>

<body>

<div align="left">
  <?php include('top.php'); ?>
<table width="852" border="0" align="center" cellpadding="1" cellspacing="2">
    <tr bgcolor="#999999">
      <td colspan="5"><div align="center"><span class="style1">Customer Invoice Generating &amp; Emailling </span></div></td>
    </tr>
    <form action="" method="post" name="Functions" target="_self" id="Functions"><tr align="center" valign="middle" bgcolor="#333333">
      <th colspan="5">
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr align="center" valign="middle">
    <td width="27%"><span class="style2">
    </span></td>
    <td width="29%">        <span class="style2">
        <input name="cmdKillAllNow" type="submit" id="cmdKillAllNow" value="Delete Current History of Process">
</span></td>
    <td width="39%">&nbsp;</td>
    <td width="2%">&nbsp;</td>
    <td width="3%">&nbsp;</td>
  </tr>
</table>
</th>
    </tr></form>
    <tr bgcolor="#333333">
      <th><div align="center"><strong>Customer ID</strong></div></th>
      <th><strong>Billing Name</strong></th>
      <th><div align="center">No. Items </div></th>
      <th><div align="center">Total Charge.</div></th>
      <th>Invoice No. </th>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#666666">
      <td width="99"><div align="center"></div></td>
      <td width="101">&nbsp;</td>
      <td width="227"><div align="center"></div></td>
      <td width="306"><div align="center"></div></td>
      <td width="97">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>
<?php

?>
