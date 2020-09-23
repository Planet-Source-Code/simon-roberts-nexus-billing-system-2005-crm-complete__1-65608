<?php
		$xmlcomplant=trim($HTTP_ACCEPT);
	
		$browser=trim($HTTP_USER_AGENT);
	
		$pagetitle = "Wap Artist #$wapartist - $profilename";
	
		$htmlxml ="xmlns=\"http://www.w3.org/1999/xhtml\"";
	
		$doctype = "html PUBLIC \"-//NOKIA//DTD XHTML Mobile +CHTML 1.0//EN\" \"http://www.nokia.com/dtd/xhtml-mp-chtml.dtd\"";
	
	  if (strpos(" " . strtoupper($xmlcomplant),"XML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"XHTML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
		
	  if (strpos(" " . strtoupper($xmlcomplant),"WML") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"MP3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}
	  
	  if (strpos(" " . strtoupper($xmlcomplant),"MPEG3") != false)			// Ericsson WAP phones and emulators
		{
			$mp3 = true;
		}	
		
	  if (strpos(" " . strtoupper($xmlcomplant),"M4A") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}

	  if (strpos(" " . strtoupper($xmlcomplant),"MP4") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}			
		
	  if (strpos(" " . strtoupper($xmlcomplant),"MPEG4") != false)			// Ericsson WAP phones and emulators
		{
			$mp4 = true;
		}											
	  if (strpos(" " . strtoupper($browser),"R380") != false)			// Ericsson WAP phones and emulators
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPI") != false)			// Ericsson WapIDE 2.0
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MC21") != false)			// Ericsson MC218
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"AUR ") != false)			// Ericsson R320
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ERIC") != false)			// Ericsson R380
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UP.B") != false)			// UP.Browser
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WINW") != false)			// WinWAP browser
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UPG1") != false)			// UP.SDK 4.0
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"UPSI") != false)			// another kind of UP.Browser ??
		{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"QWAP") != false)			// unknown QWAPPER browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"JIGS") != false)			// unknown JigSaw browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"JAVA") != false)			// unknown Java based browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ALCA") != false)			// unknown Alcatel-BE3 browser (UP based?)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MITS") != false)			// unknown Mitsubishi browser
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MOT-") != false)			// unknown browser (UP based?)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MY S") != false)           // unknown Ericsson devkit browser ?
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPJ") != false)			// Virtual WAPJAG www.wapjag.de
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"FETC") != false)			// fetchpage.cgi Perl script from www.wapcab.de
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"ALAV") != false)			// yet another unknown UP based browser ?
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"WAPA") != false)
			{
			$xmlyes = true;
		}
	  if (strpos(" " . strtoupper($browser),"MOZI") != false)
			{
			$xmlyes = true;
		} 
	  if (strpos(" " . strtoupper($browser),"NOKI") != false)             // another unknown browser (Web based "Wapalyzer"?)
		{
			$xmlyes = true;
		}
		

if ($xmlyes == true) {
?>
<?php echo sprintf("%s%sxml version=\"1.0\"? encoding=\"utf\-8\"%s%s","<","?","?",">"); ?> 
<!DOCTYPE <?php echo $doctype; ?>> 
<html <?php echo $htmlxml; ?>>
<head>
<title><?php echo $pagetitle; ?></title>
<?php include("http://www.projectalpha.com.au/wap/wapcss.php?wapartist=0&pagefrom=default"); ?>
</head>

<body>
<div align="center">
  <p><img src="../images/idents/ep_ident.gif" width="90%" height="90%"></p>
  <p class="style1">Free Ringtones<br>
    <a href="ringtones2.php"><br>
    Select Here For The New Rock Selection
  </a> </p>
  <p>Use your Download Link function in your gprs to download a midi file (*) this should be playable on a polyphonic phone.<br>
    <br>
    <strong class="style1">Movie Soundtracks</strong><br>
    <br>
    <a href="../audio/midi/aladdin.mid">aladdin</a><br>
    <a href="../audio/midi/alien.mid">alien</a><br>
    <a href="../audio/midi/back1.mid">back1</a><br>
    <a href="../audio/midi/blackhole.mid">blackhole</a><br>
    <a href="../audio/midi/bluebro.mid">bluebro</a><br>
    <a href="../audio/midi/bond.mid">bond</a><br>
    <a href="../audio/midi/bond77.mid">bond77</a><br>
    <a href="../audio/midi/devil_hm.mid">devil hm</a><br>
    <a href="../audio/midi/gbusters.mid">gbusters</a><br>
    <a href="../audio/midi/st_gen.mid">st gen</a><br>
    <a href="../audio/midi/st_khan.mid">st khan</a><br>
    <a href="../audio/midi/star_trek_6.mid">star trek 6</a><br>
    <a href="../audio/midi/star_trek_next_generation1.mid">star trek next generation1</a><br>
    <a href="../audio/midi/superman.mid">superman</a><br>
    <a href="../audio/midi/t2-gen.mid">t2-gen</a><br>
    <a href="../audio/midi/taxi.mid">taxi</a><br>
    <a href="../audio/midi/topgun.mid">topgun</a><br>
    <a href="../audio/midi/war_wrld.mid">war wrld
  </a><br>
  <a href="../audio/midi/movie_theme_songs/apollo_13.mid">apollo 13</a><br>
  <a href="../audio/midi/movie_theme_songs/batman_theme.mid">batman theme  </a><br>
  <a href="../audio/midi/movie_theme_songs/beatlejuice_theme_song.mid">beatlejuice theme song</a><br>
  <a href="../audio/midi/movie_theme_songs/forrest_gump_theme.mid">forrest gump theme  </a><br>
  <a href="../audio/midi/movie_theme_songs/ghost_busters.mid">ghost busters  </a><br>
  <a href="../audio/midi/movie_theme_songs/godfather.mid">godfather </a><br>
  <a href="../audio/midi/movie_theme_songs/gone_with_the_wind.mid">gone with the wind</a><br>
  <a href="../audio/midi/movie_theme_songs/mortal_kombat.mid">mortal kombat  </a><br>
  <a href="../audio/midi/movie_theme_songs/naked_gun_theme.mid">naked gun theme</a><br>  
  <a href="../audio/midi/movie_theme_songs/pulp_fiction.mid">pulp fiction</a><br>
  <a href="../audio/midi/movie_theme_songs/robin_hood_(eve&#352;ing_i%27d_do).mid">robin hood (eve&Scaron;ing i'd do)</a><br>
  <a href="../audio/midi/movie_theme_songs/theme_from_james_bond.mid">theme from james bond  </a><br>
  <a href="../audio/midi/movie_theme_songs/theme_from_jaws.mid">theme from jaws  </a><br>
  <a href="../audio/midi/movie_theme_songs/true_lies.mid">true lies  </a><br>
  <br>  
    <img src="../images/new.gif" width="29" height="11"><br>
    <br>
    <span class="style1"><strong>Star Wars</strong></span> </p>
  <p><a href="../audio/midi/starwars/empire.mid">empire</a><br>
    <a href="../audio/midi/starwars/empire2.mid">empire2</a><br>
    <a href="../audio/midi/starwars/esbfinal.mid">esbfinal</a><br>
    <a href="../audio/midi/starwars/esbtheme.mid">esbtheme</a><br>
    <a href="../audio/midi/starwars/ewok.mid">ewok</a><br>
    <a href="../audio/midi/starwars/ewok_2.mid">ewok 2</a><br>
  <a href="../audio/midi/starwars/hanleiag.mid">hanleiag</a> <br>
  </p>
</div>
</body>
</html>
<?php } ?>