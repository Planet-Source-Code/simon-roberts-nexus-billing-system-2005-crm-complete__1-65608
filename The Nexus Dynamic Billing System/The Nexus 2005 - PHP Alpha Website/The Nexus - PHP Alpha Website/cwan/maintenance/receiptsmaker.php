<?php require_once('../../Connections/projectalpha.php'); ?>

<?php

	function AddReceiptItem( $acci_RecID, $InvoiceOutID, $OnlineSessionSysopID, $acciServicesID , $Paid, $Refunded,  $GSTPaid, $GSTRefunded, $DateNava, $Description,$PaymentType,$SerialNumber, $RefundID, $TraxrID, $StatementID) 
	{
		$receiptsqla = "insert into receipts (`acci_RecID`,`Description`,`PaymentType`,`SerialNumber`,`RefundID`,`acciServicesID`,`TraxrID`,`InvoiceOutID`,`InvoiceInID`,`Paid`,`Refunded`,`StatementID`,`GSTRefunded`,`GSTPaid`) ";
		$receiptsqla .= "VALUES ('%d','%s','$s','$s','%d','%d','%d','%d','%d','%d','%d','%d','%d','%d')";
		$receiptsqlb = sprintf($receiptsqla, $acci_RecID, $Description, $PaymentType, $SerialNumber, $RefundID, $acciServicesID, $TraxrID, $InvoiceOutID, $InvoiceInID, $Paid, $Refunded, $StatementID, $GSTRefunded, $GSTPaid);
		   
		mysql_select_db($database_projectalpha, $projectalpha);
		$newreciept = mysql_query($receiptsqlb, $projectalpha) or die(mysql_error());
	
	}
	
?>