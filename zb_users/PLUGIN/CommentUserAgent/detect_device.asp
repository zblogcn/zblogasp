<script language="javascript" runat="server" >

// Detect Console or Mobile Device
function detect_device(useragent)
 {
    var _link = "",
    title_e = "",
	title_c = "",
    code = "",
    regmatch;
    // Apple
    if (/iPad/i.test(useragent))
    {
        _link = "http://www.apple.com/itunes";
        title_e = "iPad";

        if (/CPU\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /CPU\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "ipad";

    }
    else if (/iPod/i.test(useragent))
    {
        _link = "http://www.apple.com/itunes";
        title_e = "iPod";

        if (/iPhone\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /iPhone\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "iphone";

    }
    else if (/iPhone/i.test(useragent))
    {
        _link = "http://www.apple.com/iphone";
        title_e = "iPhone";

        if (/iPhone\ OS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /iPhone\ OS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " iOS " + regmatch[1].replace(/_/g, ".");

        }

        code = "iphone";

    }
    else if (/iOs/i.test(useragent))
    {
        _link = "http://www.apple.com/";
        title_e = "iOs";

        if (/iOS\ ([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /iOS\ ([._0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1].replace(/_/g, ".");

        }

        code = "iphone";

    }

    // BenQ-Siemens (Openwave)
    else if (/[^M]SIE/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/BenQ-Siemens";
        title_e = "BenQ-Siemens";

        if (/[^M]SIE-([.0-9a-zA-Z]+)\//i.test(useragent)) {
            regmatch = /[^M]SIE-([.0-9a-zA-Z]+)\//i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "benq-siemens";

    }
	
    // 魅族
    else if (/(MEIZU (MX|M9))/i.test(useragent))
    {
        _link = "http://www.meizu.com/";
        title_e = "meizu";
		title_c = "魅族";
        code = "meizu";

    }

	//小米
    else if (/MI-ONE|MI 2/i.test(useragent))
    {
        _link = "http://www.xiaomi.com/";
        title_e = "XiaoMi";
		title_c = "小米";
        code = "xiaomi";

    }

    // BlackBerry
    else if (/BlackBerry|PlayBook|BB10/i.test(useragent))
    {
        _link = "http://www.blackberry.com/";
        title_e = "BlackBerry";

        if (/blackberry([.0-9a-zA-Z]+)\//i.test(useragent)) {
            regmatch = /blackberry([.0-9a-zA-Z]+)\//i.exec(useragent);
            title_e += " " + regmatch[1];

        }
		else if (useragent.indexOf("BB10")>0)
		{
			title_e = "BB10";
		}

        code = "blackberry";

    }

    // Dell
    else if (/Dell Mini 5/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Dell_Streak";
        title_e = "Dell Mini 5";
        code = "dell";

    }
    else if (/Dell Streak/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Dell_Streak";
        title_e = "Dell Streak";
        code = "dell";

    }
    else if (/Dell/i.test(useragent)) {
        _link = "http://en.wikipedia.org/wiki/Dell";
        title_e = "Dell";
        code = "dell";

    }

    // Google
    else if (/Nexus One/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Nexus_One";
        title_e = "Nexus One";
        code = "google-nexusone";

    }

    // HTC
    else if (/Desire/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Desire";
        title_e = "HTC Desire";
        code = "htc";

    }
    else if (/Rhodium/i.test(useragent)
    || /HTC[_|\ ]Touch[_|\ ]Pro2/i.test(useragent)
    || /WMD-50433/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Touch_Pro2";
        title_e = "HTC Touch Pro2";
        code = "htc";

    }
    else if (/HTC[_|\ ]Touch[_|\ ]Pro/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/HTC_Touch_Pro";
        title_e = "HTC Touch Pro";
        code = "htc";

    }
    else if (/HTC/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/High_Tech_Computer_Corporation";
        title_e = "HTC";

        if (/HTC[\ |_|-]8500/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Startrek";
            title_e += " Startrek";

        }
        else if (/HTC[\ |_|-]Hero/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Hero";
            title_e += " Hero";

        }
        else if (/HTC[\ |_|-]Legend/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Legend";
            title_e += " Legend";

        }
        else if (/HTC[\ |_|-]Magic/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Magic";
            title_e += " Magic";

        }
        else if (/HTC[\ |_|-]P3450/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Touch";
            title_e += " Touch";

        }
        else if (/HTC[\ |_|-]P3650/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Polaris";
            title_e += " Polaris";

        }
        else if (/HTC[\ |_|-]S710/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_S710";
            title_e += " S710";

        }
        else if (/HTC[\ |_|-]Tattoo/i.test(useragent))
        {
            _link = "http://en.wikipedia.org/wiki/HTC_Tattoo";
            title_e += " Tattoo";

        }
        else if (/HTC[\ |_|-]?([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /HTC[\ |_|-]?([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }
        else if (/HTC([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /HTC([._0-9a-zA-Z]+)/i.exec(useragent);
            title_e += str_replace("_", " ", regmatch[1]);

        }

        code = "htc";

    }
	// huawei
	else if (/Huawei/i.test(useragent))
	{
		_link = "http://www.huawei.com/cn/";
		title_e = "HuaWei";
		title_c = "华为";
		code = "huawei";
		regmatch = /HUAWEI([.0-9a-zA-Z]+)/i.exec(useragent);
		title_e += " " + regmatch[1];
		title_c += " " + regmatch[1];
	}

    // Kindle
    else if (/Kindle/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Amazon_Kindle";
        title_e = "Kindle";

        if (/Kindle\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Kindle\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "kindle";

    }
    // Lenovo
    else if (/Lenovo/i.test(useragent))
    {
        _link = "http://www.lenovo.com.cn";
        title_e = "Lenovo";
		title_c = "联想"

        if (/Lenovo[\ |-|\/]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Lenovo[\ |-|\/]([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "lenovo";

    }
    // LG
    else if (/LG/i.test(useragent))
    {
        _link = "http://www.lgmobile.com";
        title_e = "LG";

        if (/LG[E]?[\ |-|\/]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /LG[E]?[\ |-|\/]([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "lg";

    }


    // Motorola
    else if (/\ Droid/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola_Droid";
        title_e += "Motorola Droid";
        code = "motorola";

    }
    else if (/XT720/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola";
        title_e += "Motorola Motoroi (XT720)";
        code = "motorola";

    }
    else if (/MOT-/i.test(useragent)
    || /MIB/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Motorola";
        title_e = "Motorola";

        if (/MOTO([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /MOTO([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }
        if (/MOT-([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /MOT-([.0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "motorola";

    }
    else if (/XOOM/i.test(useragent)) {
        _link = "http://en.wikipedia.org/wiki/Motorola_Xoom";
        title_e += "Motorola Xoom";
        code = "motorola";

    }

    // Nintendo
    else if (/Nintendo/i.test(useragent))
    {
        title_e = "Nintendo";

        if (/Nintendo DSi/i.test(useragent))
        {
            _link = "http://www.nintendodsi.com/";
            title_e += " DSi";
            code = "nintendodsi";

        }
        else if (/Nintendo DS/i.test(useragent))
        {
            _link = "http://www.nintendo.com/ds";
            title_e += " DS";
            code = "nintendods";

        }
        else if (/Nintendo Wii/i.test(useragent))
        {
            _link = "http://www.nintendo.com/wii";
            title_e += " Wii";
            code = "nintendowii";

        }
        else
        {
            _link = "http://www.nintendo.com/";
            code = "nintendo";

        }

    }

    // Nokia
    else if (/Nokia/i.test(useragent))
    {
		/*if (!(/S(eries)?60/i.test(useragent)) || !/Symbian/i.test(useragent))
		{*/
			_link = "http://www.nokia.com/";
			title_e = "Nokia";
			code = "nokia";
			if (/Nokia(E|N| )?([0-9]+)/i.test(useragent))
			{
				regmatch = /Nokia(E|N| )?([0-9]+)/i.exec(useragent);
				title_e += " " + regmatch[1] + regmatch[2];
			}
			else if (/Lumia ([0-9]+)/i.test(useragent))
			{
				regmatch = /Lumia ([0-9]+)/i.exec(useragent);
				title_e += " Lumia " + regmatch[1];
			}
		//}
    }

    // OLPC (One Laptop Per Child)
    else if (/OLPC/i.test(useragent))
    {
        _link = "http://www.laptop.org/";
        title_e = "OLPC (XO)";
        code = "olpc";

    }
    // 昂达
    else if (/onda/i.test(useragent))
    {
        _link = "http://http://www.onda.cn/";
        title_e = "Onda";
		title_c = "昂达"
        code = "onda";

    }

    // Palm
    else if (/\ Pixi\//i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Palm_Pixi";
        title_e = "Palm Pixi";
        code = "palm";

    }
    else if (/\ Pre\//i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Palm_Pre";
        title_e = "Palm Pre";
        code = "palm";

    }
    else if (/Palm/i.test(useragent))
    {
        _link = "http://www.palm.com/";
        title_e = "Palm";
        code = "palm";

    }
    else if (/wp-webos/i.test(useragent))
    {
        _link = "http://www.palm.com/";
        title_e = "Palm";
        code = "palm";

    }

    // Playstation
    else if (/PlayStation/i.test(useragent))
    {
        title_e = "PlayStation";

        if (/[PS|PlayStation\ ]3/i.test(useragent))
        {
            _link = "http://www.us.playstation.com/PS3";
            title_e += " 3";

        }
        else if (/[PlayStation Portable|PSP]/i.test(useragent))
        {
            _link = "http://www.us.playstation.com/PSP";
            title_e += " Portable";

        }
        else if (/[PlayStation Vita|PSVita]/i.test(useragent))
        {

            _link = "http://us.playstation.com/psvita/";
            title_e += " Vita";

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
        title_e = "Galaxy Nexus";
        code = "samsung";

    }
	else if (/GT-\D\d+ Build\/[0-9A-Z]+/i.test(useragent))
    {
        _link = "http://www.samsungmobile.com/";
        title_e = "Samsung";

        if (/(GT-\D\d+) Build\/[0-9A-Z]+?/i.test(useragent)) {
            regmatch = /(GT-\D\d+) Build\/[0-9A-Z]+/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "samsung";

    }
    else if (/SmartTV/i.test(useragent))
    {
        _link = "http://www.freethetvchallenge.com/details/faq";
        title_e = "Samsung Smart TV";
        code = "samsung";

    }
    else if (/Samsung/i.test(useragent))
    {
        _link = "http://www.samsungmobile.com/";
        title_e = "Samsung";

        if (/Samsung-([.\-0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Samsung-([.\-0-9a-zA-Z]+)/i.exec(useragent);
            title_e += " " + regmatch[1];

        }

        code = "samsung";

    }

    // Sony Ericsson
    else if (/SonyEricsson/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/SonyEricsson";
        title_e = "SonyEricsson";

        if (/SonyEricsson([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /SonyEricsson([.0-9a-zA-Z]+)/i.exec(useragent);
            if (regmatch[1].toLowerCase() == "u20i")
            {
                title_e += " Xperia X10 Mini Pro";

            }
            else
            {
                title_e += " " + regmatch[1];

            }

        }

        code = "sonyericsson";

    }

    // Windows Phone
    else if (/Windows Phone( OS)? [0-9.]+/.test(useragent))
	{
		var regmatch = /Windows Phone( OS)? ([0-9.]+)/.exec(useragent);
		title_e = "Windows Phone " + regmatch[2];
		code = "windowsphone";
		_link = "http://www.windowsphone.com/zh-cn/"
	}
	
	//中兴
    else if (/zte/i.test(useragent))
    {
        _link = "http://www.zte.com.cn/cn/";
        title_e = "ZTE";
		title_c = "中兴"
        code = "ZTE";

    }
	//Some special UA..
	//is MSIE
	if(/MSIE.+?Windows.+?Trident/.test(useragent) && !/Windows ?Phone/.test(useragent)){
		_link = "";
		title_e = "";
		code = "";		
	}

	title_c = title_c==""?title_e:title_c;
	
	var json = {
        "link": _link,
        "text": title_e,
        "filename": code,
        "folder": "device",
		"fullfilename":"16/device/"+code+".png",
		"ver":""


    }
    return json;

}

</script> 