<?php
# FileName="Connection_php_mysql.htm"
# Type="MYSQL"
# HTTP="true"
$hostname_projectalpha = "demon.comcen.com.au";
$database_projectalpha = "projectalpha";
$username_projectalpha = "sroberts";
$password_projectalpha = "kl7jb0lsf";
$projectalpha = mysql_pconnect($hostname_projectalpha, $username_projectalpha, $password_projectalpha) or trigger_error(mysql_error(),E_USER_ERROR); 
?>