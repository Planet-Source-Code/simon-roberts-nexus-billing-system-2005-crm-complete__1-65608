<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title><?php
			echo Sprintf('Refreshing to %s',$nURL);
			?></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<?php
			echo Sprintf('Refreshing to <a href="%s">%s</a> click if this does not occur',$nURL,$nURL);
			?>
<div align="center">

	  <script language="vbscript">
		self.location=<?php 
			echo sprintf('"%s"',$nURL); ?>
	</script>
	
    </div>
</body>
</html>
