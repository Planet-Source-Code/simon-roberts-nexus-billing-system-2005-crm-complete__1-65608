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
  <p>&nbsp;</p>
  <p class="style1"><img src="../images/rockpop.gif" width="100%" height="90%"><br>
  Ring Tones 2 </p>
  <p><br>
      <span class="style1">  Cake</span>
</p>
  <p class="style1">Pink Floyd</p>
  <p><a href="../audio/midi/pink_floyd/another_brick_in_the_wall.mid">Another Brick Against The Wall </a><br>
      <a href="../audio/midi/pink_floyd/brain_damage.mid">Brain Damage </a><br>
      <a href="../audio/midi/pink_floyd/comfortably_numb.mid">Comfortably Numb </a><br>
      <a href="../audio/midi/pink_floyd/goodbye_blue_sky.mid">Goodbye Blue Sky</a><br>
      <a href="../audio/midi/pink_floyd/hey_you.mid">Hey You</a><br>
      <a href="../audio/midi/pink_floyd/high_hopes.mid">Hight Hopes </a><br>
      <a href="../audio/midi/pink_floyd/keep_talking.mid">Keep Talking </a> <br>
      <a href="../audio/midi/pink_floyd/money.mid">Money</a><br>
      <a href="../audio/midi/pink_floyd/shine_on_you_crazy_diamond.mid">Shine On You Crazy Diamond </a><br>
      <a href="../audio/midi/pink_floyd/the_great_gig_in_the_sky.mid">The Great Gig In The Sky </a><br>
      <a href="../audio/midi/pink_floyd/the_trial.mid">The Trial </a><br>
      <a href="../audio/midi/pink_floyd/wish_you_were_here.mid">Wish You Where Here </a></p>
  <p class="style1">Red Hot Chill Peppers </p>
  <p><a href="../audio/midi/red_hot_chilli_peppers/aeroplane.mid">areoplane</a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/deep_kick.mid">deep kick</a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/higher_ground.mid">higher ground</a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/knock_me_down.mid">knock me down</a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/my_friends.mid">my friends</a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/roller_coaster_of_love.mid">roller coaster of love </a><br>
      <a href="../audio/midi/red_hot_chilli_peppers/under_the_bridge.mid">under the bridge</a></p>
  <p><br>
      <br>
      <span class="style1">Rage Against The Machine <br>
      <br>
      </span><a href="../audio/midi/rage_against_the_machine/bulls_on_parade.mid">bulls on parade </a><br>
      <a href="../audio/midi/rage_against_the_machine/killing_in_the_name_of.mid">killing in the name of</a><br>
      <a href="../audio/midi/rage_against_the_machine/know_your_enemy.mid">know your enemy</a><br>
      <a href="../audio/midi/rage_against_the_machine/people_of_the_sun.mid">people of the sun </a><br>
      <a href="../audio/midi/rage_against_the_machine/retire_me.mid">retire me</a><br>
      <a href="../audio/midi/rage_against_the_machine/revolver.mid">revolver</a><br>
      <a href="../audio/midi/rage_against_the_machine/take_the_power_back.mid">take the power back</a><br>
      <a href="../audio/midi/rage_against_the_machine/wind_below.mid">wind below</a><br>
      <a href="../audio/midi/rage_against_the_machine/year_of_tha_boomerang.mid">year of tha boomerang </a></p>
  <p>&nbsp;</p>
  <p><span class="style1">Bob Marley <br>
        <br>
    </span><a href="../audio/midi/marley_bob/i_shot_the_sheriff.mid">i shot the sheriff </a><br>
    <a href="../audio/midi/marley_bob/iron_lion_zion.mid">iron lion zion</a><br>
    <a href="../audio/midi/marley_bob/jammin.mid">jammin</a><br>
    <a href="../audio/midi/marley_bob/no_woman_no_cry.mid">no woman no cry </a><br>
    <a href="../audio/midi/marley_bob/waiting_in_vain.mid">waiting in vain </a></p>
  <p>&nbsp; </p>
  <p class="style1">Blues Brothers </p>
  <p><a href="../audio/midi/blues_brothers/everybody_needs_somebody.mid">everybody needs somebody </a><br>
      <a href="../audio/midi/blues_brothers/peter_gunn.mid">peter gunn </a><br>
      <a href="../audio/midi/blues_brothers/soul_man.mid">soul man</a><br>
      <a href="../audio/midi/blues_brothers/sweet_home_chicago.mid">sweet home chicago</a></p>
  <p>&nbsp; </p>
  <p><span class="style1">Others</span><br>
      <br>
      <a href="../audio/midi/beastie_boys/slowride.mid">beastie boys - slow ride </a><br>
      <br>
  </p>
  <p><span class="style1"><a href="../audio/midi/cake/Frank%20Sinatra.mid">Frank Sinatra</a><br>
      <a href="../audio/midi/cake/Never%20There.mid">Never There</a><br>
      <a href="../audio/midi/cake/The%20Distance.mid">The Distance</a></span><br>
  </p>
  <p><span class="style1"><br>
    <br>
    Coolio</span><br>
    <br>
    <span class="style1"><a href="../audio/midi/coolio/C%20U%20When%20U%20Get%20There.mid">C U When U Get There</a><br>
    <a href="../audio/midi/coolio/Fantastic%20Voyage.mid">Fantastic Voyage</a><br>
    <a href="../audio/midi/coolio/Gangsta%27s%20Paradise.mid">Gangsta's Paradise</a><br>
    <a href="../audio/midi/coolio/Too%20Hot.mid">Too Hot
  </a></span> </p>
  <p class="style1"><br>
    <br>
  Duran Duran </p>
  <p class="style1"><a href="../audio/midi/duran_duran/A%20View%20To%20Kill.mid">A View To Kil</a><br>
    <a href="../audio/midi/duran_duran/Chauffer.mid">Chauffer</a><br>
  <a href="../audio/midi/duran_duran/Come%20Undone.mid">Come Undone</a><br>  
  <a href="../audio/midi/duran_duran/Election%20Day.mid">Election Day</a><br>
  <a href="../audio/midi/duran_duran/Equan%20For%20You.mid">Equan For You</a><br>
  <a href="../audio/midi/duran_duran/New%20Moon%20On%20Monday.mid">New Moon On Monday</a><br>
   <a href="../audio/midi/duran_duran/%20Notorious.mid">Notorious</a><br>
   <a href="../audio/midi/duran_duran/Ordinary%20World.mid">Ordinary World</a><br>
   <a href="../audio/midi/duran_duran/Reflex.mid">Reflex</a><br>
   <a href="../audio/midi/duran_duran/Rio.mid">Rio</a><br>
   <a href="../audio/midi/duran_duran/Save%20A%20Prayer.mid">Save A Prayer</a><br>
   <br>
   <br>
   <br>
   Faith No More<br>
   <br> 
   <a href="../audio/midi/faith_no_more/Easy.mid">Easy</a><br>
   <a href="../audio/midi/faith_no_more/Epic.mid">Epic</a><br>
   <a href="../audio/midi/faith_no_more/Midnight%20Cowboy.mid">Midnight Cowboy</a><br>
   <br>
   <br>
   <br>
   Greenday
   <br>
   <br>
   <a href="../audio/midi/green_day/86.mid">86</a><br>
   <a href="../audio/midi/green_day/Basket%20Case.mid">Basket Case</a><br>
   <a href="../audio/midi/green_day/Brain%20Stew.mid">Brain Stew</a><br>
   <a href="../audio/midi/green_day/Brat.mid">Brat   </a><br>
   <a href="../audio/midi/green_day/Christie%20Road.mid">Christie Road</a><br>
   <a href="../audio/midi/green_day/Coming%20Clean.mid">Coming Clean</a><br>   
   <a href="../audio/midi/green_day/Do%20Da%20Da.mid">Do Da Da</a><br>
   <a href="../audio/midi/green_day/Emenius%20Sleepus.mid">Emenius Sleepus</a><br>
   <a href="../audio/midi/green_day/Geek%20Stink%20Breath.mid">Geek Stink Breath</a><br>
   <a href="../audio/midi/green_day/Good%20Riddance.mid">Good Riddance</a><br>
   <a href="../audio/midi/green_day/In%20The%20End.mid">In The End</a><br>
   <a href="../audio/midi/green_day/Jar.mid">Jar</a><br>
   <a href="../audio/midi/green_day/Light%20Years%20Away.mid">Light Years Away</a><br>
   <a href="../audio/midi/green_day/Longview.mid">Longview</a><br>
   <a href="../audio/midi/green_day/My%20Generation.mid">My Generation</a><br>
   <a href="../audio/midi/green_day/Only%20Of%20You.mid">Only Of You</a><br>
   <a href="../audio/midi/green_day/Pulling%20Teeth.mid">Pulling Teeth</a><br>
   <a href="../audio/midi/green_day/Reject.mid">Reject</a><br>
   <a href="../audio/midi/green_day/She.mid">She</a><br>
   <a href="../audio/midi/green_day/Stuck%20With%20Me.mid">Stuck With Me</a><br>   
   <a href="../audio/midi/green_day/The%20Grouch.mid">The Grouch</a><br>
   <a href="../audio/midi/green_day/Walking%20Contradiction.mid">Walking Contradiction</a><br>
   <a href="../audio/midi/green_day/When%20I%20Come%20Around.mid">When I Come Around</a><br>
  <a href="../audio/midi/green_day/Worry%20Rock.mid">Worry Rock</a>   </p>
  <p class="style1">&nbsp;</p>
  <p class="style1">Jamiroquai<br>
    <br>
    <a href="../audio/midi/jamiroquai/Alright.mid">Alright</a><br>
    <a href="../audio/midi/jamiroquai/Blow%20Your%20Mind.mid">Blow Your Mind</a><br>
    <a href="../audio/midi/jamiroquai/Cosmic%20Girl.mid">Cosmic Girl</a><br>
    <a href="../audio/midi/jamiroquai/Deeper%20Underground.mid">Deeper Underground</a><br>
    <a href="../audio/midi/jamiroquai/Half%20The%20Man.mid">Half The Man</a><br>
    <a href="../audio/midi/jamiroquai/Music%20Of%20The%20Mind.mid">Music Of The Mind</a><br>
    <a href="../audio/midi/jamiroquai/Too%20Young%20To%20Die.mid">Too Young To Die</a><br>
    <a href="../audio/midi/jamiroquai/Virtual%20Insanity.mid">Virtual Insanity</a><br>
    <a href="../audio/midi/jamiroquai/When%20You%20Gonna%20Learn.mid">When You Gonna Learn    </a><br>
    <br>
    <br>
    <br>
    No Doubt<br>
    <br>
    <a href="../audio/midi/no_doubt/Dont%20Speak%20Clueless%20Remix.mid">Dont Speak Clueless Remix</a><br>
    <a href="../audio/midi/no_doubt/Hey%20You.mid">Hey You</a><br>
    <a href="../audio/midi/no_doubt/Just%20A%20Girl.mid">Just A Girl</a><br>
    <a href="../audio/midi/no_doubt/Spiderwebs.mid">Spiderwebs</a>    <br>
    <br>
    <br>
    <br>
    Primus<br>
    <br>
    <a href="../audio/midi/primus/Codding%20Town.mid">Codding Town</a><br>
    <a href="../audio/midi/primus/Jerry%20Was%20A%20Racecar%20Driver.mid">Jerry Was A Racecar Driver</a><br>
    <a href="../audio/midi/primus/Precipitation.mid">Precipitation</a><br>
    <a href="../audio/midi/primus/Wynona%27s%20Big%20Brown%20Beaver.mid">Wynona's Big Brown Beaver    </a><br>
    <br>
    <br>
    <br>
    <br>
  </p>
</div>
</body>
</html>
<?php } ?>