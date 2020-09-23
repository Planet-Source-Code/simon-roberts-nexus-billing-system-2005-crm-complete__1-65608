<?php
# FileName="Connection_php_mysql.htm"
# Type="MYSQL"
# HTTP="true"
$hostname_Epwebdev = "localhost";
$database_Epwebdev = "pa_online";
$username_Epwebdev = "";
$password_Epwebdev = "";
$Epwebdev = mysql_pconnect($hostname_Epwebdev, $username_Epwebdev, $password_Epwebdev) or trigger_error(mysql_error(),E_USER_ERROR); 
?>