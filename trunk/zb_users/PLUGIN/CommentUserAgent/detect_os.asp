<script language="javascript" runat="server" >// Detect Operating System
function detect_os(useragent)
 {
    var _link = "",
    title = "",
    code = "",
	version = "",
    regmatch;

    if (/AmigaOS/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/AmigaOS";
        title = "AmigaOS";

        if (/AmigaOS\ ([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /AmigaOS\ ([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "amigaos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Android/i.test(useragent))
    {
        _link = "http://www.android.com/";
        title = "Android";
        code = "android";

        if (/Android[\ |\/]?([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Android[\ |\/]?([.0-9a-zA-Z]+)/i.exec(useragent);
            version = regmatch[1];
            title += " " + version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/[^A-Za-z]Arch/i.test(useragent))
    {
        _link = "http://www.archlinux.org/";
        title = "Arch Linux";
        code = "archlinux";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/BeOS/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/BeOS";
        title = "BeOS";
        code = "beos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/CentOS/i.test(useragent))
    {
        _link = "http://www.centos.org/";
        title = "CentOS";

        if (/.el([.0-9a-zA-Z]+).centos/i.test(useragent)) {
            regmatch = /.el([.0-9a-zA-Z]+).centos/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "centos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Chakra/i.test(useragent))
    {
        _link = "http://www.chakra-linux.org/";
        title = "Chakra Linux";
        code = "chakra";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/CrOS/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Google_Chrome_OS";
        title = "Google Chrome OS";
        code = "chromeos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Crunchbang/i.test(useragent))
    {
        _link = "http://www.crunchbanglinux.org/";
        title = "Crunchbang";
        code = "crunchbang";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Debian/i.test(useragent))
    {
        _link = "http://www.debian.org/";
        title = "Debian GNU/Linux";
        code = "debian";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/DragonFly/i.test(useragent))
    {
        _link = "http://www.dragonflybsd.org/";
        title = "DragonFly BSD";
        code = "dragonflybsd";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Edubuntu/i.test(useragent))
    {
        _link = "http://www.edubuntu.org/";
        title = "Edubuntu";

        if (/Edubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Edubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (regmatch[1] < 10)
        {
            code = "edubuntu-1";

        }
        else
        {
            code = "edubuntu-2";

        }

        if (version.length > 1)
        {
            title += version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Fedora/i.test(useragent))
    {
        _link = "http://www.fedoraproject.org/";
        title = "Fedora";

        if (/.fc([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /.fc([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "fedora";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Foresight\ Linux/i.test(useragent))
    {
        _link = "http://www.foresightlinux.org/";
        title = "Foresight Linux";

        if (/Foresight\ Linux\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Foresight\ Linux\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "foresight";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/FreeBSD/i.test(useragent))
    {
        _link = "http://www.freebsd.org/";
        title = "FreeBSD";
        code = "freebsd";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Gentoo/i.test(useragent))
    {
        _link = "http://www.gentoo.org/";
        title = "Gentoo";
        code = "gentoo";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Inferno/i.test(useragent))
    {
        _link = "http://www.vitanuova.com/inferno/";
        title = "Inferno";
        code = "inferno";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/IRIX/i.test(useragent))
    {
        _link = "http://www+sgi.com/partners/?/technology/irix/";
        title = "IRIX Linux";

        if (/IRIX(64)?\ ([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /IRIX(64)?\ ([.0-9a-zA-Z]+)/i.exec(useragent);
            if (regmatch[1])
            {
                title += " x" + regmatch[1];

            }
            if (regmatch[2])
            {
                title += " " + regmatch[2];

            }

        }

        code = "irix";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Kanotix/i.test(useragent))
    {
        _link = "http://www.kanotix.com/";
        title = "Kanotix";
        code = "kanotix";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Knoppix/i.test(useragent))
    {
        _link = "http://www.knoppix.net/";
        title = "Knoppix";
        code = "knoppix";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Kubuntu/i.test(useragent))
    {
        _link = "http://www.kubuntu.org/";
        title = "Kubuntu";

        if (/Kubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Kubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (regmatch[1] < 10)
        {
            code = "kubuntu-1";

        }
        else
        {
            code = "kubuntu-2";

        }

        if (version.length > 1)
        {
            title += version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/LindowsOS/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Lsongs";
        title = "LindowsOS";
        code = "lindowsos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Linspire/i.test(useragent))
    {
        _link = "http://www.linspire.com/";
        title = "Linspire";
        code = "lindowsos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Linux\ Mint/i.test(useragent))
    {
        _link = "http://www.linuxmint.com/";
        title = "Linux Mint";

        if (/Linux\ Mint\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Linux\ Mint\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "linuxmint";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Lubuntu/i.test(useragent))
    {
        _link = "http://www.lubuntu.net/";
        title = "Lubuntu";

        if (/Lubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Lubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (regmatch[1] < 10)
        {
            code = "lubuntu-1";

        }
        else
        {
            code = "lubuntu-2";

        }

        if (version.length > 1)
        {
            title += version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Mageia/i.test(useragent))
    {
        _link = "http://www.mageia.org/";
        title = "Mageia";
        code = "mageia";

    }
    else if (/Mandriva/i.test(useragent))
    {
        _link = "http://www.mandriva.com/";
        title = "Mandriva";

        if (/mdv([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /mdv([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "mandriva";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/moonOS/i.test(useragent))
    {
        _link = "http://www.moonos.org/";
        title = "moonOS";

        if (/moonOS\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /moonOS\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "moonos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/MorphOS/i.test(useragent))
    {
        _link = "http://www.morphos-team.net/";
        title = "MorphOS";
        code = "morphos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/NetBSD/i.test(useragent))
    {
        _link = "http://www.netbsd.org/";
        title = "NetBSD";
        code = "netbsd";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Nova/i.test(useragent))
    {
        _link = "http://www.nova.cu";
        title = "Nova";

        if (/Nova[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Nova[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (version.length > 1)
        {
            title += version;

        }

        code = "nova";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/OpenBSD/i.test(useragent))
    {
        _link = "http://www.openbsd.org/";
        title = "OpenBSD";
        code = "openbsd";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Oracle/i.test(useragent))
    {
        _link = "http://www.oracle+com/us/technologies/linux/";
        title = "Oracle";

        if (/.el([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /.el([._0-9a-zA-Z]+)/i.exec(useragent);
            title += " Enterprise Linux "+regmatch[1].replace(/_/g,".");

        }
        else
        {
            title += " Linux";

        }

        code = "oracle";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Pardus/i.test(useragent))
    {
        _link = "http://www.pardus.org.tr/en/";
        title = "Pardus";
        code = "pardus";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/PCLinuxOS/i.test(useragent))
    {
        _link = "http://www.pclinuxos.com/";
        title = "PCLinuxOS";

        if (/PCLinuxOS\/[.\-0-9a-zA-Z]+pclos([.\-0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /PCLinuxOS\/[.\-0-9a-zA-Z]+pclos([.\-0-9a-zA-Z]+)/i.exec(useragent);
            title += " "+regmatch[1].replace(/_/g,".");

        }

        code = "pclinuxos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Red\ Hat/i.test(useragent)
    || /RedHat/i.test(useragent))
    {
        _link = "http://www.redhat.com/";
        title = "Red Hat";

        if (/.el([._0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /.el([._0-9a-zA-Z]+)/i.exec(useragent);
            title += " Enterprise Linux "+regmatch[1].replace(/_/g,".");

        }

        code = "red-hat";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Rosa/i.test(useragent))
    {
        _link = "http://www.rosalab.com/";
        title = "Rosa Linux";
        code = "rosa";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Sabayon/i.test(useragent))
    {
        _link = "http://www+sabayonlinux.org/";
        title = "Sabayon Linux";
        code = "sabayon";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Slackware/i.test(useragent))
    {
        _link = "http://www+slackware.com/";
        title = "Slackware";
        code = "slackware";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Solaris/i.test(useragent))
    {
        _link = "http://www+sun.com/software/solaris/";
        title = "Solaris";
        code = "solaris";

    }
    else if (/SunOS/i.test(useragent))
    {
        _link = "http://www+sun.com/software/solaris/";
        title = "Solaris";
        code = "solaris";

    }
    else if (/Suse/i.test(useragent))
    {
        _link = "http://www.opensuse.org/";
        title = "openSUSE";
        code = "suse";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Symb[ian]?[OS]?/i.test(useragent))
    {
        _link = "http://www+symbianos.org/";
        title = "SymbianOS";

        if (/Symb[ian]?[OS]?\/([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Symb[ian]?[OS]?\/([.0-9a-zA-Z]+)/i.exec(useragent);
            title += " " + regmatch[1];

        }

        code = "symbianos";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Unix/i.test(useragent))
    {
        _link = "http://www.unix.org/";
        title = "Unix";
        code = "unix";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/VectorLinux/i.test(useragent))
    {
        _link = "http://www.vectorlinux.com/";
        title = "VectorLinux";
        code = "vectorlinux";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Venenux/i.test(useragent))
    {
        _link = "http://www.venenux.org/";
        title = "Venenux GNU Linux";
        code = "venenux";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/webOS/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/WebOS";
        title = "Palm webOS";
        code = "palm";

    }
    else if (/Windows/i.test(useragent)
    || /WinNT/i.test(useragent)
    || /Win32/i.test(useragent))
    {
        _link = "http://www.microsoft.com/windows/";

        if (/Windows NT 6.2; Win64; x64;/i.test(useragent)
        || /Windows NT 6.2; WOW64/i.test(useragent))
        {
            title = "Windows 8 x64 Edition";
            code = "win-5";

        }
        else if (/Windows NT 6.2/i.test(useragent))
        {
            title = "Windows 8";
            code = "win-5";

        }
        else if (/Windows NT 6.1; Win64; x64;/i.test(useragent)
        || /Windows NT 6.1; WOW64/i.test(useragent))
        {
            title = "Windows 7 x64 Edition";
            code = "win-4";

        }
        else if (/Windows NT 6.1/i.test(useragent))
        {
            title = "Windows 7";
            code = "win-4";

        }
        else if (/Windows NT 6.0/i.test(useragent))
        {
            title = "Windows Vista";
            code = "win-3";

        }
        else if (/Windows NT 5.2 x64/i.test(useragent))
        {
            title = "Windows XP x64 Edition";
            code = "win-2";

        }
        else if (/Windows NT 5.2; Win64; x64;/i.test(useragent))
        {
            title = "Windows Server 2003 x64 Edition";
            code = "win-2";

        }
        else if (/Windows NT 5.2/i.test(useragent))
        {
            title = "Windows Server 2003";
            code = "win-2";

        }
        else if (/Windows NT 5.1/i.test(useragent)
        || /Windows XP/i.test(useragent))
        {
            title = "Windows XP";
            code = "win-2";

        }
        else if (/Windows NT 5.01/i.test(useragent))
        {
            title = "Windows 2000, Service Pack 1 (SP1)";
            code = "win-1";

        }
        else if (/Windows NT 5.0/i.test(useragent)
        || /Windows 2000/i.test(useragent))
        {
            title = "Windows 2000";
            code = "win-1";

        }
        else if (/Windows NT 4.0/i.test(useragent)
        || /WinNT4.0/i.test(useragent))
        {
            title = "Microsoft Windows NT 4.0";
            code = "win-1";

        }
        else if (/Windows NT 3.51/i.test(useragent)
        || /WinNT3.51/i.test(useragent))
        {
            title = "Microsoft Windows NT 3.11";
            code = "win-1";

        }
        else if (/Windows 3.11/i.test(useragent)
        || /Win3.11/i.test(useragent)
        || /Win16/i.test(useragent))
        {
            title = "Microsoft Windows 3.11";
            code = "win-1";

        }
        else if (/Windows 3.1/i.test(useragent))
        {
            title = "Microsoft Windows 3.1";
            code = "win-1";

        }
        else if (/Windows 98; Win 9x 4.90/i.test(useragent)
        || /Win 9x 4.90/i.test(useragent)
        || /Windows ME/i.test(useragent))
        {
            title = "Windows Millennium Edition (Windows Me)";
            code = "win-1";

        }
        else if (/Win98/i.test(useragent))
        {
            title = "Windows 98 SE";
            code = "win-1";

        }
        else if (/Windows 98/i.test(useragent)
        || /Windows\ 4.10/i.test(useragent))
        {
            title = "Windows 98";
            code = "win-1";

        }
        else if (/Windows 95/i.test(useragent)
        || /Win95/i.test(useragent))
        {
            title = "Windows 95";
            code = "win-1";

        }
        else if (/Windows CE/i.test(useragent))
        {
            title = "Windows CE";
            code = "win-2";

        }
        else if (/WM5/i.test(useragent))
        {
            title = "Windows Mobile 5";
            code = "win-phone";

        }
        else if (/WindowsMobile/i.test(useragent))
        {
            title = "Windows Mobile";
            code = "win-phone";

        }
        else
        {
            title = "Windows";
            code = "win-2";

        }

    }
    else if (/Xandros/i.test(useragent))
    {
        _link = "http://www.xandros.com/";
        title = "Xandros";
        code = "xandros";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Xubuntu/i.test(useragent))
    {
        _link = "http://www.xubuntu.org/";
        title = "Xubuntu";

        if (/Xubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Xubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (regmatch[1] < 10)
        {
            code = "xubuntu-1";

        }
        else
        {
            code = "xubuntu-2";

        }

        if (version.length > 1)
        {
            title += version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Zenwalk/i.test(useragent))
    {
        _link = "http://www.zenwalk.org/";
        title = "Zenwalk GNU Linux";
        code = "zenwalk";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }

    // Pulled out of order to help ensure better detection for above platforms
    else if (/Ubuntu/i.test(useragent))
    {
        _link = "http://www.ubuntu.com/";
        title = "Ubuntu";

        if (/Ubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /Ubuntu[\/|\ ]([.0-9a-zA-Z]+)/i.exec(useragent);
            version += " " + regmatch[1];

        }

        if (regmatch[1] < 10)
        {
            code = "ubuntu-1";

        }
        else
        {
            code = "ubuntu-2";

        }

        if (version.length > 1)
        {
            title += version;

        }

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/Linux/i.test(useragent))
    {
        _link = "http://www.linux.org/";
        title = "GNU/Linux";
        code = "linux";

        if (/x86_64/i.test(useragent))
        {
            title += " x64";

        }

    }
    else if (/J2ME\/MIDP/i.test(useragent))
    {
        _link = "http://java+sun.com/javame/";
        title = "J2ME/MIDP Device";
        code = "java";

    }
	
	
	//I don't know that why some browsers' useragent have MACOSX although it run under Windows.
	else if (/Mac/i.test(useragent)
    || /Darwin/i.test(useragent))
    {
        _link = "http://www.apple+com/macosx/";

        if (/Mac OS X/i.test(useragent))
        {
			
			
            title = useragent.substr(0,useragent.toString().toLowerCase().indexOf("mac os x"));
            title = title.substr(0, title.indexOf(")"));

            if (title.indexOf(";")>=0)
            {
                title = title.substr(0, title.indexOf(";"));

            }

            title = title.replace(/_/g, ".");
            code = "mac-3";

        }
        else if (/Mac OSX/i.test(useragent))
        {
            title = useragent.substr(0,useragent.toLowerCase().indexOf("Mac OS X".toLowerCase()));
            title = title.substr(0, title.indexOf(")"));

            if (title.indexOf(";")>=0)
            {
                title = title.substr(0, title.indexOf(";"));

            }

            title = title.replace(/_/g, ".");
            code = "mac-2";

        }
        else if (/Darwin/i.test(useragent))
        {
            title = "Mac OS Darwin";
            code = "mac-1";

        }
        else
        {
            title = "Macintosh";
            code = "mac-1";

        }

    }


    // How should we display this?
	var json = {
        "link": _link,
        "text": title,
        "filename": code,
        "folder": "os",
		"ver":version,
		"fullfilename":"16/os/"+code+".png"

    }
    return json;


}
</script>