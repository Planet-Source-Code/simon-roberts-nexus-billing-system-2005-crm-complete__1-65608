<?php

$hostname = gethostbyaddr($_SERVER['REMOTE_ADDR']);
$hosts = gethostbynamel('www.example.com');
$protocol = 'tcp';
$get_prot = getprotobyname($protocol);

echo 'hostname: ' . $hostname;
echo 'hosts: ' . $hosts;
echo 'protocol: ' . $hosts;

if ($get_prot == -1) {
   // if nothing found, returns -1
   echo 'Invalid Protocol';
} else {
   echo 'Protocol #' . $get_prot;
}

$services = array('http', 'ftp', 'ssh', 'telnet', 'imap', 
'smtp', 'nicname', 'gopher', 'finger', 'pop3', 'www');

foreach ($services as $service) {                    
   $port = getservbyname($service, 'tcp');
   echo $service . ": " . $port . "<br />\n";
}

?>