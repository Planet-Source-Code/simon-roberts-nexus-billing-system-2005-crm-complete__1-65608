<?php

function &getMailer()
{
	
	include_once "xoopsmailer.php";
	
	return new XoopsMailer();
}
?>