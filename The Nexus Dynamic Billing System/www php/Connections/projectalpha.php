<?php
# FileName="Connection_php_mysql.htm"
# Type="MYSQL"
# HTTP="true"
$hostname_projectalpha = "localhost";
$database_projectalpha = "pa_online";
$username_projectalpha = "";
$password_projectalpha = "";
$projectalpha = mysql_pconnect($hostname_projectalpha, $username_projectalpha, $password_projectalpha) or trigger_error(mysql_error(),E_USER_ERROR); 



?>