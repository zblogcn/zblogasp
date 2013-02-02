<script language="javascript" runat="server">

// Detect Console or Mobile Device
function detect_device(useragent)
 {
    var _link = "",
    title = "",
    code = "",
    regmatch;
    // Apple
    if (/iPad/i.test(useragent))
    {
        _link = "http://www.apple.com/itunes";
        title = "iPad";

        if (/CPU\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /CPU\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "ipad";

    }
    else if (/iPod/i.test(useragent))
    {
        _link = "http://www.apple.com/itunes";
        title = "iPod";

        if (/iPhone\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /iPhone\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "iphone";

    }
    else if (/iPhone/i.test(useragent))
    {
        _link = "http://www.apple.com/iphone";
        title = "iPhone";

        if (/iPhone\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /iPhone\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "iphone";

    }

    // BenQ-Siemens (Openwave)
    else if (/[^M]SIE/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/BenQ-Siemens";
        title = "BenQ-Siemens";

        if (/[^M]SIE-([.0-9a-zA-Z]+)\//i.test(useragent)) {
            regmatch = /[^M]SIE-([.0-9a-zA-Z]+)\//i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "benq-siemens";

    }
	
    // 魅族
    else if (/(MEIZU (MX|M9))/i.test(useragent))
    {
        _link = "http://www.meizu.com/";
        title = "meizu";
        code = "meizu";

    }

	//小米
    else if (/MI-ONE/i.test(useragent))
    {
        _link = "http://www.xiaomi.com/";
        title = "XiaoMi";
        code = "xiaomi";

    }

    // BlackBerry
    else if (/BlackBerry|PlayBook|BB10/i.test(useragent))
    {
        _link = "http://www.blackberry.com/";
        title = "BlackBerry";

        if (/blackberry([.0-9a-zA-Z]+)\//i.test(useragent)) {
            regmatch = /blackberry([.0-9a-zA-Z]+)\//i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "blackberry";

    }

    // Dell
    else if (/Dell Mini 5/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Dell_Streak";
        title = "Dell Mini 5";
        code = "dell";

    }
    else if (/Dell Streak/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Dell_Streak";
        title = "Dell Streak";
        code = "dell";

    }
    else if (/Dell/i.test(useragent)) {
        _link = "http://en.wikipedia.org/wiki/Dell";
        title = "Dell";
        code = "dell";

    }

    // Google
    else if (/Nexus One/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Nexus_One";
        title = "Nexus One";
        code = "google-nexusone";

    }

    // HTC
    else if (/Desire/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Desire";
        title = "HTC Desire";
        code = "htc";

    }
    else if (/Rhodium/i.test(useragent)
    || /HTC[_|\ ]Touch[_|\ ]Pro2/i.test(useragent)
    || /WMD-50433/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Touch_Pro2";
        title = "HTC Touch Pro2";
        code = "htc";

    }
    else if (/HTC[_|\ ]Touch[_|\ ]Pro/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Touch_Pro";
        title = "HTC Touch Pro";
        code = "htc";

    }
    else if (/HTC/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/High_Tech_Computer_Corporation";
        title = "HTC";

        if (/HTC[\ |_|-]8500/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Startrek";
            title += " Startrek";

        }
        else if (/HTC[\ |_|-]Hero/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Hero";
            title += " Hero";

        }
        else if (/HTC[\ |_|-]Legend/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Legend";
            title += " Legend";

        }
        else if (/HTC[\ |_|-]Magic/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Magic";
            title += " Magic";

        }
        else if (/HTC[\ |_|-]P3450/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Touch";
            title += " Touch";

        }
        else if (/HTC[\ |_|-]P3650/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Polaris";
            title += " Polaris";

        }
        else if (/HTC[\ |_|-]S710/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_S710";
            title += " S710";

        }
        else if (/HTC[\ |_|-]Tattoo/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Tattoo";
            title += " Tattoo";

        }
        else if (/HTC[\ |_|-]?([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /HTC[\ |_|-]?([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }
        else if (/HTC([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /HTC([._0-9a-zA-Z]+)/i.exec(useragent);
            title += str_replace("_", " ", regmatch[1]);

        }

        code = "htc";

    }

    // Kindle
    else if (/Kindle/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Amazon_Kindle";
        title = "Kindle";

        if (/Kindle\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Kindle\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "kindle";

    }

    // LG
    else if (/LG/i.test(useragent))
    {
        _link = "http://www.lgmobile.com";
        title = "LG";

        if (/LG[E]?[\ |-|\/]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /LG[E]?[\ |-|\/]([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "lg";

    }

    // Microsoft
    else if (/Windows Phone OS 7.0/i.test(useragent)
    || /ZuneWP7/i.test(useragent)
    || /WP7/i.test(useragent))
    {
        _link = "http://www.microsoft.com/windowsphone/";
        title += "Windows Phone 7";
        code = "windowsphone";

    }

    // Motorola
    else if (/\ Droid/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola_Droid";
        title += "Motorola Droid";
        code = "motorola";

    }
    else if (/XT720/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola";
        title += "Motorola Motoroi (XT720)";
        code = "motorola";

    }
    else if (/MOT-/i.test(useragent)
    || /MIB/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola";
        title = "Motorola";

        if (/MOTO([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /MOTO([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }
        if (/MOT-([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /MOT-([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "motorola";

    }
    else if (/XOOM/i.test(useragent)) {
        _link = "http://en.wikipedia.org/wiki/Motorola_Xoom";
        title += "Motorola Xoom";
        code = "motorola";

    }

    // Nintendo
    else if (/Nintendo/i.test(useragent))
    {
        title = "Nintendo";

        if (/Nintendo DSi/i.test(useragent))
        {
            _link = "http://www.nintendodsi.com/";
            title += " DSi";
            code = "nintendodsi";

        }
        else if (/Nintendo DS/i.test(useragent))
        {
            _link = "http://www.nintendo.com/ds";
            title += " DS";
            code = "nintendods";

        }
        else if (/Nintendo Wii/i.test(useragent))
        {
            _link = "http://www.nintendo.com/wii";
            title += " Wii";
            code = "nintendowii";

        }
        else
        {
            _link = "http://www.nintendo.com/";
            code = "nintendo";

        }

    }

    // Nokia
    else if (/Nokia/i.test(useragent) && !(/S(eries)?60/i.test(useragent)))
    {
        _link = "http://www.nokia.com/";
        title = "Nokia";
        if (/Nokia(E|N)?([0-9]+)/i.test(useragent)) {
            regmatch = /Nokia(E|N)?([0-9]+)/i.exec(useragent);
            _link = "http://www.s60.com/";
            title = "Nokia Series60";
            code = "nokia";
        }

    }

    // OLPC (One Laptop Per Child)
    else if (/OLPC/i.test(useragent))
    {
        _link = "http://www.laptop.org/";
        title = "OLPC (XO)";
        code = "olpc";

    }

    // Palm
    else if (/\ Pixi\//i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Palm_Pixi";
        title = "Palm Pixi";
        code = "palm";

    }
    else if (/\ Pre\//i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Palm_Pre";
        title = "Palm Pre";
        code = "palm";

    }
    else if (/Palm/i.test(useragent))
    {
        _link = "http://www.palm.com/";
        title = "Palm";
        code = "palm";

    }
    else if (/wp-webos/i.test(useragent))
    {
        _link = "http://www.palm.com/";
        title = "Palm";
        code = "palm";

    }

    // Playstation
    else if (/PlayStation/i.test(useragent))
    {
        title = "PlayStation";

        if (/[PS|PlayStation\ ]3/i.test(useragent))
        {
            _link = "http://www.us.playstation.com/PS3";
            title += " 3";

        }
        else if (/[PlayStation Portable|PSP]/i.test(useragent))
        {
            _link = "http://www.us.playstation.com/PSP";
            title += " Portable";

        }
        else if (/[PlayStation Vita|PSVita]/i.test(useragent))
        {

            _link = "http://us.playstation.com/psvita/";
            title += " Vita";

        }
        else
        {
            _link = "http://www.us.playstation.com/";

        }

        code = "playstation";

    }

    // Samsung
    else if (/Galaxy Nexus/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Galaxy_Nexus";
        title = "Galaxy Nexus";
        code = "samsung";

    }
    else if (/SmartTV/i.test(useragent))
    {
        _link = "http://www.freethetvchallenge.com/details/faq";
        title = "Samsung Smart TV";
        code = "samsung";

    }
    else if (/Samsung/i.test(useragent))
    {
        _link = "http://www.samsungmobile.com/";
        title = "Samsung";

        if (/Samsung-([.\-0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Samsung-([.\-0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "samsung";

    }

    // Sony Ericsson
    else if (/SonyEricsson/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/SonyEricsson";
        title = "SonyEricsson";

        if (/SonyEricsson([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /SonyEricsson([.0-9a-zA-Z]+)/i.exec(useragent);
            if (regmatch[1].toLowerCase() == "u20i")
            {
                title += " Xperia X10 Mini Pro";

            }
            else
            {
                title += " " + regmatch[1];

            }

        }

        code = "sonyericsson";

    }

    // Windows Phone
    else if (/wp-windowsphone/i.test(useragent))
    {
        _link = "http://www.windowsphone.com/";
        title = "Windows Phone";
        code = "windowsphone";

    }

	//Some special UA..
	//is MSIE
	if(/MSIE.+?Windows.+?Trident/.test(useragent)){
		_link = "";
		title = "";
		code = "";		
	}


	var json = {
        "link": _link,
        "text": title,
        "filename": code,
        "folder": "device",
		"fullfilename":"16/device/"+code+".png",
		"ver":""


    }
    return json;

}

</script> 