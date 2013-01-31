<script language="javascript" runat="server">

// Detect Web Browser versions
function detect_browser_ver(title, useragent)
 {
	var json={
		"version":"",
		"full":""
	}
    // Fix for Opera's UA string changes in v10+00+ (and others)
    var start = title;
    if ((title.toLowerCase() == "Opera".toLowerCase()
    || title.toLowerCase() == "Opera Next".toLowerCase()
    || title.toLowerCase() == "Opera Labs".toLowerCase())
    && /Version/i.test(useragent))
    {
        start = "Version";

    }
    else if (title.toLowerCase() == "Opera Mobi".toLowerCase()
    && /Version/i.test(useragent))
    {
        start = "Version";

    }
    else if (title.toLowerCase() == "Safari".toLowerCase()
    && /Version/i.test(useragent))
    {
        start = "Version";

    }
    else if (title.toLowerCase() == "Pre".toLowerCase()
    && /Version/i.test(useragent))
    {
        start = "Version";

    }
    else if (title.toLowerCase() == "Android Webkit".toLowerCase())
    {
        start = "Version";

    }
    else if (title.toLowerCase() == "Links".toLowerCase())
    {
        start = "Links (";

    }
    else if (title.toLowerCase() == "UC Browser".toLowerCase())
    {
        start = "UC Browse";

    }
    else if (title.toLowerCase() == "TenFourFox".toLowerCase())
    {
        start = " rv";

    }
    else if (title.toLowerCase() == "Classilla".toLowerCase())
    {
        start = " rv";

    }
    else if (title.toLowerCase() == "SmartTV".toLowerCase())
    {
        start = "WebBrowser";

    }

    // Grab the browser version if its present
    var regmatch = new RegExp(start + '[\ |\/]?([\+0-9a-zA-Z\.]+)', "i");
	
	if(regmatch.test(useragent)){
		regmatch = regmatch.exec(useragent)
	
		version = regmatch[1];
		
	
		// json.full=browser Title and Version, but first++some titles need to be changed
		if (title.toLowerCase() == "msie"
		&& version.toLowerCase() == "7.0"
		&& /Trident\/4|5|6+0/i.test(useragent))
		{
			json.full=" 8.0+ (Compatibility Mode)";
			// Fix for IE8 quirky UA string with Compatibility Mode enabled
	
		}
		else if (title.toLowerCase() == "msie")
		{
			json.full=" " + version;
	
		}
		else if (title.toLowerCase() == "multi-browser")
		{
			json.full="Multi-Browser XP " + version;
	
		}
		else if (title.toLowerCase() == "nf-browser")
		{
			json.full="NetFront " + version;
	
		}
		else if (title.toLowerCase() == "semc-browser")
		{
			json.full="SEMC Browser " + version;
	
		}
		else if (title.toLowerCase() == "ucweb")
		{
			json.full="UC Browser " + version;
	
		}
		else if (title.toLowerCase() == "up.browser"
		|| title.toLowerCase() == "up.link")
		{
			json.full="Openwave Mobile Browser " + version;
	
		}
		else if (title.toLowerCase() == "chromeframe")
		{
			json.full="Google Chrome Frame " + version;
	
		}
		else if (title.toLowerCase() == "mozilladeveloperpreview")
		{
			json.full="Mozilla Developer Preview " + version;
	
		}
		else if (title.toLowerCase() == "multi-browser")
		{
			json.full="Multi-Browser XP " + version;
	
		}
		else if (title.toLowerCase() == "opera mobi")
		{
			json.full="Opera Mobile " + version;
	
		}
		else if (title.toLowerCase() == "osb-browser")
		{
			json.full="Gtk+ WebCore " + version;
	
		}
		else if (title.toLowerCase() == "tablet browser")
		{
			json.full="MicroB " + version;
	
		}
		else if (title.toLowerCase() == "tencenttraveler")
		{
			json.full="Tencent Traveler " + version;
	
		}
		else if (title.toLowerCase() == "crmo")
		{
			json.full="Chrome Mobile " + version;
	
		}
		else if (title.toLowerCase() == "smarttv")
		{
			json.full="Maple Browser " + version;
	
		}
		else if (title.toLowerCase() == "wp-android"
		|| title.toLowerCase() == "wp-iphone")
		{
			//TODO check into Android version being returned
			json.full="Wordpress App " + version;
	
		}
		else if (title.toLowerCase() == "atomicbrowser")
		{
			json.full="Atomic Web Browser " + version;
	
		}
		else if (title.toLowerCase() == "barcapro")
		{
			json.full="Barca Pro " + version;
	
		}
		else if (title.toLowerCase() == "dplus")
		{
			json.full="D+ " + version;
	
		}
		else if (title.toLowerCase() == "opera labs")
		{
			var regmatch = /Edition\ Labs([\ +_0-9a-zA-Z]+);/i.exec(useragent);
			json.full=title + regmatch[1] + " " + version;
	
		}
		else
		{
			json.full=title + " " + version;
	
		}
		json.version=version;
	}
	return json;

}
</script>