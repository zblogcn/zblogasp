<script language="javascript" runat="server">
// Detect Web Browsers
function detect_webbrowser(useragent)
 {
    var _link = "",
    title = "",
    code = "",
    ver = "",
    regmatch;


    var mobile = 0;

    if (/360se|360ee/i.test(useragent))
    {
        _link = "http://se.360.cn/";
        title = "360UnSafe Explorer";
        code = "360se";



    }
    else if (/Abolimba/i.test(useragent))
    {
        _link = "http://www.aborange.de/products/freeware/abolimba-multibrowser.php";
        title = "Abolimba";
        code = "abolimba";



    }
    else if (/Acoo\ Browser/i.test(useragent))
    {
        _link = "http://www.acoobrowser.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Acoo " + _ver.full;
        ver = _ver.version;
        code = "acoobrowser";



    }
    else if (/Alienforce/i.test(useragent))
    {
        _link = "http://sourceforge.net/projects/alienforce/";
        _ver = detect_browser_ver("Alienforce", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "alienforce";



    }
    else if (/Amaya/i.test(useragent))
    {
        _link = "http://www.w3.org/Amaya/";
        _ver = detect_browser_ver("Amaya", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "amaya";



    }
    else if (/Amiga-AWeb/i.test(useragent))
    {
        _link = "http://aweb.sunsite.dk/";
        _ver = detect_browser_ver("AWeb", useragent);
        title = "Amiga " + _ver.full;
        ver = _ver.version;
        code = "amiga-aweb";



    }
    else if (/America\ Online\ Browser/i.test(useragent))
    {
        _link = "http://downloads.channel.aol.com/browser";
        _ver = detect_browser_ver("Browser", useragent);
        title = "America Online " + _ver.full;
        ver = _ver.version;
        code = "aol";



    }
    else if (/AmigaVoyager/i.test(useragent))
    {
        _link = "http://v3.vapor.com/voyager/";
        _ver = detect_browser_ver("Voyager", useragent);
        title = "Amiga " + _ver.full;
        ver = _ver.version;
        code = "amigavoyager";



    }
    else if (/AOL/i.test(useragent))
    {
        _link = "http://downloads.channel.aol.com/browser";
        _ver = detect_browser_ver("AOL", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "aol";



    }
    else if (/Arora/i.test(useragent))
    {
        _link = "http://code.google.com/p/arora/";
        _ver = detect_browser_ver("Arora", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "arora";



    }
    else if (/AtomicBrowser/i.test(useragent))
    {
        _link = "http://www.atomicwebbrowser.com/";
        _ver = detect_browser_ver("AtomicBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "atomicwebbrowser";



    }
    else if (/Avant\ (Browser|TriCore)/i.test(useragent))
    {
        _link = "http://www.avantbrowser.com/";
        title = "Avant Browser";
        code = "avantbrowser";



    }
    else if (/ba?idubrowser/i.test(useragent))
    {
        _link = "http://liulanqi.baidu.com/";
        _ver = detect_browser_ver("BIDUBrowser", useragent)
		if(_ver.version=="") _ver=detect_browser_ver("BAIDUBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "bidubrowser";



    }
    else if (/BarcaPro/i.test(useragent))
    {
        _link = "http://www.pocosystems.com/home/index.php?option=content&task=category&sectionid=2&id=9&Itemid=27";
        _ver = detect_browser_ver("BarcaPro", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "barca";



    }
    else if (/Barca/i.test(useragent))
    {
        _link = "http://www.pocosystems.com/home/index.php?option=content&task=category&sectionid=2&id=9&Itemid=27";
        _ver = detect_browser_ver("Barca", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "barca";



    }
    else if (/Beamrise/i.test(useragent))
    {
        _link = "http://www.beamrise.com/";
        _ver = detect_browser_ver("Beamrise", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "beamrise";



    }
    else if (/Beonex/i.test(useragent))
    {
        _link = "http://www.beonex.com/";
        _ver = detect_browser_ver("Beonex", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "beonex";



    }
    else if (/BlackBerry/i.test(useragent))
    {
        _link = "http://www.blackberry.com/";
        _ver = detect_browser_ver("BlackBerry", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "blackberry";



    }
    else if (/Blackbird/i.test(useragent))
    {
        _link = "http://www.blackbirdbrowser.com/";
        _ver = detect_browser_ver("Blackbird", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "blackbird";



    }
    else if (/BlackHawk/i.test(useragent))
    {
        _link = "http://www.netgate.sk/blackhawk/help/welcome-to-blackhawk-web-browser.html";
        _ver = detect_browser_ver("BlackHawk", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "blackhawk";



    }
    else if (/Blazer/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Blazer_(web_browser)";
        _ver = detect_browser_ver("Blazer", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "blazer";



    }
    else if (/Bolt/i.test(useragent))
    {
        _link = "http://www.boltbrowser.com/";
        _ver = detect_browser_ver("Bolt", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "bolt";



    }
    else if (/BonEcho/i.test(useragent))
    {
        _link = "http://www.mozilla.org/projects/minefield/";
        _ver = detect_browser_ver("BonEcho", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/BrowseX/i.test(useragent))
    {
        _link = "http://pdqi.com/browsex/";
        title = "BrowseX";
        code = "browsex";



    }
    else if (/Browzar/i.test(useragent))
    {
        _link = "http://www.browzar.com/";
        _ver = detect_browser_ver("Browzar", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "browzar";



    }
    else if (/Bunjalloo/i.test(useragent))
    {
        _link = "http://code.google.com/p/quirkysoft/";
        _ver = detect_browser_ver("Bunjalloo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "bunjalloo";



    }
    else if (/Camino/i.test(useragent))
    {
        _link = "http://www.caminobrowser.org/";
        _ver = detect_browser_ver("Camino", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "camino";



    }
    else if (/Cayman\ Browser/i.test(useragent))
    {
        _link = "http://www.caymanbrowser.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Cayman " + _ver.full;
        ver = _ver.version;
        code = "caymanbrowser";



    }
    else if (/Charon/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Charon_(web_browser)";
        _ver = detect_browser_ver("Charon", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/Cheshire/i.test(useragent))
    {
        _link = "http://downloads.channel.aol.com/browser";
        _ver = detect_browser_ver("Cheshire", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "aol";



    }
    else if (/Chimera/i.test(useragent))
    {
        _link = "http://www.chimera.org/";
        _ver = detect_browser_ver("Chimera", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/chromeframe/i.test(useragent))
    {
        _link = "http://code.google.com/chrome/chromeframe/";
        _ver = detect_browser_ver("chromeframe", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "google";



    }
    else if (/ChromePlus/i.test(useragent))
    {
        _link = "http://www.chromeplus.org/";
        _ver = detect_browser_ver("ChromePlus", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "chromeplus";



    }
    else if (/Iron/i.test(useragent))
    {
        _link = "http://www.srware.net/";
        _ver = detect_browser_ver("Iron", useragent);
        title = "SRWare " + _ver.full;
        ver = _ver.version;
        code = "srwareiron";



    }
    else if (/Chromium/i.test(useragent))
    {
        _link = "http://www.chromium.org/";
        _ver = detect_browser_ver("Chromium", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "chromium";



    }
    else if (/Classilla/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Classilla";
        _ver = detect_browser_ver("Classilla", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "classilla";



    }
    else if (/Columbus/i.test(useragent))
    {
        _link = "http://www.columbus-browser.com/";
        _ver = detect_browser_ver("Columbus", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "columbus";



    }
    else if (/CometBird/i.test(useragent))
    {
        _link = "http://www.cometbird.com/";
        _ver = detect_browser_ver("CometBird", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "cometbird";



    }
    else if (/Comodo_Dragon/i.test(useragent))
    {
        _link = "http://www.comodo.com/home/internet-security/browser.php";
        _ver = detect_browser_ver("Dragon", useragent);
        title = "Comodo " + _ver.full;
        ver = _ver.version;
        code = "comodo-dragon";



    }
    else if (/Conkeror/i.test(useragent))
    {
        _link = "http://www.conkeror.org/";
        _ver = detect_browser_ver("Conkeror", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "conkeror";



    }
    else if (/CoolNovo/i.test(useragent))
    {
        _link = "http://www.coolnovo.com/";
        _ver = detect_browser_ver("CoolNovo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "coolnovo";



    }
    else if (/Crazy\ Browser/i.test(useragent))
    {
        _link = "http://www.crazybrowser.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Crazy " + _ver.full;
        ver = _ver.version;
        code = "crazybrowser";



    }
    else if (/CrMo/i.test(useragent))
    {
        _link = "http://www.google.com/chrome";
        _ver = detect_browser_ver("CrMo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "chrome";



    }
    else if (/Cruz/i.test(useragent))
    {
        _link = "http://www.cruzapp.com/";
        _ver = detect_browser_ver("Cruz", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "cruz";



    }
    else if (/Cyberdog/i.test(useragent))
    {
        _link = "http://www.cyberdog.org/about/cyberdog/cyberbrowse.html";
        _ver = detect_browser_ver("Cyberdog", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "cyberdog";



    }
    else if (/DPlus/i.test(useragent))
    {
        _link = "http://dplus-browser.sourceforge.net/";
        _ver = detect_browser_ver("DPlus", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "dillo";



    }
    else if (/Deepnet\ Explorer/i.test(useragent))
    {
        _link = "http://www.deepnetexplorer.com/";
        _ver = detect_browser_ver("Deepnet Explorer", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "deepnetexplorer";



    }
    else if (/Demeter/i.test(useragent))
    {
        _link = "http://www.hurrikenux.com/Demeter/";
        _ver = detect_browser_ver("Demeter", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "demeter";



    }
    else if (/DeskBrowse/i.test(useragent))
    {
        _link = "http://www.deskbrowse.org/";
        _ver = detect_browser_ver("DeskBrowse", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "deskbrowse";



    }
    else if (/Dillo/i.test(useragent))
    {
        _link = "http://www.dillo.org/";
        _ver = detect_browser_ver("Dillo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "dillo";



    }
    else if (/DoCoMo/i.test(useragent))
    {
        _link = "http://www.nttdocomo.com/";
        _ver = detect_browser_ver("DoCoMo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/DocZilla/i.test(useragent))
    {
        _link = "http://www.doczilla.com/";
        _ver = detect_browser_ver("DocZilla", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "doczilla";



    }
    else if (/Dolfin/i.test(useragent))
    {
        _link = "http://www.samsungmobile.com/";
        _ver = detect_browser_ver("Dolfin", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "samsung";



    }
    else if (/Dooble/i.test(useragent))
    {
        _link = "http://dooble.sourceforge.net/";
        _ver = detect_browser_ver("Dooble", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "dooble";



    }
    else if (/Doris/i.test(useragent))
    {
        _link = "http://www.anygraaf.fi/browser/indexe.htm";
        _ver = detect_browser_ver("Doris", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "doris";



    }
    else if (/Dorothy/i.test(useragent))
    {
        _link = "http://www.dorothybrowser.com/";
        _ver = detect_browser_ver("Dorothy", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "dorothybrowser";



    }
    else if (/Edbrowse/i.test(useragent))
    {
        _link = "http://edbrowse.sourceforge.net/";
        _ver = detect_browser_ver("Edbrowse", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "edbrowse";



    }
    else if (/E_links/i.test(useragent))
    {
        _link = "http://e_links.or.cz/";
        _ver = detect_browser_ver("E_links", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "e_links";



    }
    else if (/Element\ Browser/i.test(useragent))
    {
        _link = "http://www.elementsoftware.co.uk/software/elementbrowser/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Element " + _ver.full;
        ver = _ver.version;
        code = "elementbrowser";



    }
    else if (/Enigma\ Browser/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Enigma_Browser";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Enigma " + _ver.full;
        ver = _ver.version;
        code = "enigmabrowser";



    }
    else if (/EnigmaFox/i.test(useragent))
    {
        _link = "#";
        _ver = detect_browser_ver("EnigmaFox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/Epic/i.test(useragent))
    {
        _link = "http://www.epicbrowser.com/";
        _ver = detect_browser_ver("Epic", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "epicbrowser";



    }
    else if (/Epiphany/i.test(useragent))
    {
        _link = "http://gnome.org/projects/epiphany/";
        _ver = detect_browser_ver("Epiphany", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "epiphany";



    }
    else if (/Escape/i.test(useragent))
    {
        _link = "http://www.espial.com/products/evo_browser/";
        _ver = detect_browser_ver("Escape", useragent);
        title = "Espial TV Browser - " + _ver.full;
        ver = _ver.version;
        code = "espialtvbrowser";



    }
    else if (/Fennec/i.test(useragent))
    {
        _link = "https://wiki.mozilla.org/Fennec";
        _ver = detect_browser_ver("Fennec", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "fennec";



    }
    else if (/Firebird/i.test(useragent))
    {
        _link = "http://seb.mozdev.org/firebird/";
        _ver = detect_browser_ver("Firebird", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firebird";



    }
    else if (/Fireweb\ Navigator/i.test(useragent))
    {
        _link = "http://www.arsslensoft.tk/?q=node/7";
        _ver = detect_browser_ver("Fireweb Navigator", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firewebnavigator";



    }
    else if (/Flock/i.test(useragent))
    {
        _link = "http://www.flock.com/";
        _ver = detect_browser_ver("Flock", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "flock";



    }
    else if (/Fluid/i.test(useragent))
    {
        _link = "http://www.fluidapp.com/";
        _ver = detect_browser_ver("Fluid", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "fluid";



    }
    else if (/Galaxy/i.test(useragent)
    && !/Chrome/i.test(useragent))
    {
        _link = "http://www.traos.org/";
        _ver = detect_browser_ver("Galaxy", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "galaxy";



    }
    else if (/Galeon/i.test(useragent))
    {
        _link = "http://galeon.sourceforge.net/";
        _ver = detect_browser_ver("Galeon", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "galeon";



    }
    else if (/GlobalMojo/i.test(useragent))
    {
        _link = "http://www.globalmojo.com/";
        _ver = detect_browser_ver("GlobalMojo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "globalmojo";



    }
    else if (/GoBrowser/i.test(useragent))
    {
        _link = "http://www.gobrowser.cn/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "GO " + _ver.full;
        ver = _ver.version;
        code = "gobrowser";



    }

    else if (/Google\ Wireless\ Transcoder/i.test(useragent))
    {
        _link = "http://google.com/gwt/n";
        title = "Google Wireless Transcoder";
        code = "google";



    }
    else if (/GoSurf/i.test(useragent))
    {
        _link = "http://gosurfbrowser.com/?ln=en";
        _ver = detect_browser_ver("GoSurf", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "gosurf";



    }
    else if (/GranParadiso/i.test(useragent))
    {
        _link = "http://www.mozilla.org/";
        _ver = detect_browser_ver("GranParadiso", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/GreenBrowser/i.test(useragent))
    {
        _link = "http://www.morequick.com/";
        _ver = detect_browser_ver("GreenBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "greenbrowser";



    }
    else if (/Hana/i.test(useragent))
    {
        _link = "http://www.alloutsoftware.com/";
        _ver = detect_browser_ver("Hana", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "hana";



    }
    else if (/HotJava/i.test(useragent))
    {
        _link = "http://java.sun.com/products/archive/hotjava/";
        _ver = detect_browser_ver("HotJava", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "hotjava";



    }
    else if (/Hv3/i.test(useragent))
    {
        _link = "http://tkhtml.tcl.tk/hv3.html";
        _ver = detect_browser_ver("Hv3", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "hv3";



    }
    else if (/Hydra\ Browser/i.test(useragent))
    {
        _link = "http://www.hydrabrowser.com/";
        title = "Hydra Browser";
        code = "hydrabrowser";



    }
    else if (/Iris/i.test(useragent))
    {
        _link = "http://www.torchmobile.com/";
        _ver = detect_browser_ver("Iris", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "iris";



    }
    else if (/IBM\ WebExplorer/i.test(useragent))
    {
        _link = "http://www.networking.ibm.com/WebExplorer/";
        _ver = detect_browser_ver("WebExplorer", useragent);
        title = "IBM " + _ver.full;
        ver = _ver.version;
        code = "ibmwebexplorer";



    }
    else if (/IBrowse/i.test(useragent))
    {
        _link = "http://www.ibrowse-dev.net/";
        _ver = detect_browser_ver("IBrowse", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ibrowse";



    }
    else if (/iCab/i.test(useragent))
    {
        _link = "http://www.icab.de/";
        _ver = detect_browser_ver("iCab", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "icab";



    }
    else if (/Ice Browser/i.test(useragent))
    {
        _link = "http://www.icesoft.com/products/icebrowser.html";
        _ver = detect_browser_ver("Ice Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "icebrowser";



    }
    else if (/Iceape/i.test(useragent))
    {
        _link = "http://packages.debian.org/iceape";
        _ver = detect_browser_ver("Iceape", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "iceape";



    }
    else if (/IceCat/i.test(useragent))
    {
        _link = "http://gnuzilla.gnu.org/";
        _ver = detect_browser_ver("IceCat", useragent);
        title = "GNU " + _ver.full;
        ver = _ver.version;
        code = "icecat";



    }
    else if (/IceWeasel/i.test(useragent))
    {
        _link = "http://www.geticeweasel.org/";
        _ver = detect_browser_ver("IceWeasel", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "iceweasel";



    }
    else if (/IEMobile/i.test(useragent))
    {
        _link = "http://www.microsoft.com/windowsmobile/en-us/downloads/microsoft/internet-explorer-mobile.mspx";
        _ver = detect_browser_ver("IEMobile", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "msie-mobile";



    }
    else if (/iNet\ Browser/i.test(useragent))
    {
        _link = "http://alexanderjbeston.wordpress.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "iNet " + _ver.full;
        ver = _ver.version;
        code = "null";



    }
    else if (/iRider/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/IRider";
        _ver = detect_browser_ver("iRider", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "irider";



    }
    else if (/Iron/i.test(useragent))
    {
        _link = "http://www.srware.net/en/software_srware_iron.php";
        _ver = detect_browser_ver("Iron", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "iron";



    }
    else if (/InternetSurfboard/i.test(useragent))
    {
        _link = "http://inetsurfboard.sourceforge.net/";
        _ver = detect_browser_ver("InternetSurfboard", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "internetsurfboard";



    }
    else if (/Jasmine/i.test(useragent))
    {
        _link = "http://www.samsungmobile.com/";
        _ver = detect_browser_ver("Jasmine", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "samsung";



    }
    else if (/K-Meleon/i.test(useragent))
    {
        _link = "http://kmeleon.sourceforge.net/";
        _ver = detect_browser_ver("K-Meleon", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kmeleon";



    }
    else if (/K-Ninja/i.test(useragent))
    {
        _link = "http://k-ninja-samurai.en.softonic.com/";
        _ver = detect_browser_ver("K-Ninja", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kninja";



    }
    else if (/Kapiko/i.test(useragent))
    {
        _link = "http://ufoxlab.googlepages.com/cooperation";
        _ver = detect_browser_ver("Kapiko", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kapiko";



    }
    else if (/Kazehakase/i.test(useragent))
    {
        _link = "http://kazehakase.sourceforge.jp/";
        _ver = detect_browser_ver("Kazehakase", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kazehakase";



    }
    else if (/Strata/i.test(useragent))
    {
        _link = "http://www.kirix.com/";
        _ver = detect_browser_ver("Strata", useragent);
        title = "Kirix " + _ver.full;
        ver = _ver.version;
        code = "kirix-strata";



    }
    else if (/KKman/i.test(useragent))
    {
        _link = "http://www.kkman.com.tw/";
        _ver = detect_browser_ver("KKman", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kkman";



    }
    else if (/KMail/i.test(useragent))
    {
        _link = "http://kontact.kde.org/kmail/";
        _ver = detect_browser_ver("KMail", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kmail";



    }
    else if (/KMLite/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/K-Meleon";
        _ver = detect_browser_ver("KMLite", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kmeleon";



    }
    else if (/Konqueror/i.test(useragent))
    {
        _link = "http://konqueror.kde.org/";
        _ver = detect_browser_ver("Konqueror", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "konqueror";



    }
    else if (/Kylo/i.test(useragent))
    {
        _link = "http://kylo.tv/";
        _ver = detect_browser_ver("Kylo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "kylo";



    }
    else if (/LBrowser/i.test(useragent))
    {
        _link = "http://wiki.freespire.org/index.php/Web_Browser";
        _ver = detect_browser_ver("LBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "lbrowser";



    }
    else if (/LBBrowser/i.test(useragent))
    {
        _link = "http://www.liebao.cn/";
        title = "LieBao";
        code = "liebao";



    }

    else if (/LeechCraft/i.test(useragent))
    {
        _link = "http://leechcraft.org/";
        title = "LeechCraft";
        code = "null";



    }
    else if (/_links/i.test(useragent)
    && !/online\ _link\ validator/i.test(useragent))
    {
        _link = "http://_links.sourceforge.net/";
        _ver = detect_browser_ver("_links", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "_links";



    }
    else if (/Lobo/i.test(useragent))
    {
        _link = "http://www.lobobrowser.org/";
        _ver = detect_browser_ver("Lobo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "lobo";



    }
    else if (/lolifox/i.test(useragent))
    {
        _link = "http://www.lolifox.com/";
        _ver = detect_browser_ver("lolifox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "lolifox";



    }
    else if (/Lorentz/i.test(useragent))
    {
        _link = "http://news.softpedia.com/news/Firefox-Codenamed-Lorentz-Drops-in-March-2010-130855.shtml";
        _ver = detect_browser_ver("Lorentz", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/Lunascape/i.test(useragent))
    {
        _link = "http://www.lunascape.tv";
        _ver = detect_browser_ver("Lunascape", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "lunascape";



    }
    else if (/Lynx/i.test(useragent))
    {
        _link = "http://lynx.browser.org/";
        _ver = detect_browser_ver("Lynx", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "lynx";



    }
    else if (/Madfox/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Madfox";
        _ver = detect_browser_ver("Madfox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "madfox";



    }
    else if (/Maemo\ Browser/i.test(useragent))
    {
        _link = "http://maemo.nokia.com/features/maemo-browser/";
        _ver = detect_browser_ver("Maemo Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "maemo";



    }
    else if (/Maxthon/i.test(useragent))
    {
        _link = "http://www.maxthon.com/";
        _ver = detect_browser_ver("Maxthon", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "maxthon";



    }
    else if (/\ MIB\//i.test(useragent))
    {
        _link = "http://www.motorola.com/content.jsp?globalObjectId=1827-4343";
        _ver = detect_browser_ver("MIB", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "mib";



    }
    else if (/Tablet\ browser/i.test(useragent))
    {
        _link = "http://browser.garage.maemo.org/";
        _ver = detect_browser_ver("Tablet browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "microb";



    }
    else if (/Midori/i.test(useragent))
    {
        _link = "http://www.twotoasts.de/index.php?/pages/midori_summary.html";
        _ver = detect_browser_ver("Midori", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "midori";



    }
    else if (/Minefield/i.test(useragent))
    {
        _link = "http://www.mozilla.org/projects/minefield/";
        _ver = detect_browser_ver("Minefield", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "minefield";



    }
    else if (/MiniBrowser/i.test(useragent))
    {
        _link = "http://dmkho.tripod.com/";
        _ver = detect_browser_ver("MiniBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "minibrowser";



    }
    else if (/Minimo/i.test(useragent))
    {
        _link = "http://www-archive.mozilla.org/projects/minimo/";
        _ver = detect_browser_ver("Minimo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "minimo";



    }
    else if (/Mosaic/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Mosaic_(web_browser)";
        _ver = detect_browser_ver("Mosaic", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "mosaic";



    }
    else if (/MozillaDeveloperPreview/i.test(useragent))
    {
        _link = "http://www.mozilla.org/projects/devpreview/releasenotes/";
        _ver = detect_browser_ver("MozillaDeveloperPreview", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/MQQBrowser/i.test(useragent))
    {
        _link = "http://browser.qq.com/";
		_ver = detect_browser_ver("MQQBrowser", useragent);
		ver = _ver.version;
        title = _ver.full;
        code = "qqbrowser";



    }
	else if (/QQBrowser/i.test(useragent))
    {
        _link = "http://browser.qq.com/"; 
		_ver = detect_browser_ver("QQBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "qqbrowser";



    }
    else if (/Multi-Browser/i.test(useragent))
    {
        _link = "http://www.multibrowser.de/";
        _ver = detect_browser_ver("Multi-Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "multi-browserxp";



    }
    else if (/MultiZilla/i.test(useragent))
    {
        _link = "http://multizilla.mozdev.org/";
        _ver = detect_browser_ver("MultiZilla", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "mozilla";



    }
    else if (/myibrow/i.test(useragent)
    && /My\ Internet\ Browser/i.test(useragent))
    {
        _link = "http://myinternetbrowser.webove-stranky.org/";
        _ver = detect_browser_ver("myibrow", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "my-internet-browser";



    }
    else if (/MyIE2/i.test(useragent))
    {
        _link = "http://www.myie2.com/";
        _ver = detect_browser_ver("MyIE2", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "myie2";



    }
    else if (/Namoroka/i.test(useragent))
    {
        _link = "https://wiki.mozilla.org/Firefox/Namoroka";
        _ver = detect_browser_ver("Namoroka", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/Navigator/i.test(useragent))
    {
        _link = "http://netscape.aol.com/";
        _ver = detect_browser_ver("Navigator", useragent);
        title = "Netscape " + _ver.full;
        ver = _ver.version;
        code = "netscape";



    }
    else if (/NetBox/i.test(useragent))
    {
        _link = "http://www.netgem.com/";
        _ver = detect_browser_ver("NetBox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netbox";



    }
    else if (/NetCaptor/i.test(useragent))
    {
        _link = "http://www.netcaptor.com/";
        _ver = detect_browser_ver("NetCaptor", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netcaptor";



    }
    else if (/NetFront/i.test(useragent))
    {
        _link = "http://www.access-company.com/";
        _ver = detect_browser_ver("NetFront", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netfront";



    }
    else if (/NetNewsWire/i.test(useragent))
    {
        _link = "http://www.newsgator.com/individuals/netnewswire/";
        _ver = detect_browser_ver("NetNewsWire", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netnewswire";



    }
    else if (/NetPositive/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/NetPositive";
        _ver = detect_browser_ver("NetPositive", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netpositive";



    }
    else if (/Netscape/i.test(useragent))
    {
        _link = "http://netscape.aol.com/";
        _ver = detect_browser_ver("Netscape", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netscape";



    }
    else if (/NetSurf/i.test(useragent))
    {
        _link = "http://www.netsurf-browser.org/";
        _ver = detect_browser_ver("NetSurf", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netsurf";



    }
    else if (/NF-Browser/i.test(useragent))
    {
        _link = "http://www.access-company.com/";
        _ver = detect_browser_ver("NF-Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "netfront";



    }
    else if (/NokiaBrowser/i.test(useragent))
    {
        _link = "http://browser.nokia.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Nokia " + _ver.full;
        ver = _ver.version;
        code = "nokia";



    }
    else if (/Novarra-Vision/i.test(useragent))
    {
        _link = "http://www.novarra.com/";
        _ver = detect_browser_ver("Vision", useragent);
        title = "Novarra " + _ver.full;
        ver = _ver.version;
        code = "novarra";



    }
    else if (/Obigo/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Obigo_Browser";
        _ver = detect_browser_ver("Obigo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "obigo";



    }
    else if (/OffByOne/i.test(useragent))
    {
        _link = "http://www.offbyone.com/";
        title = "Off By One";
        code = "offbyone";



    }
    else if (/OmniWeb/i.test(useragent))
    {
        _link = "http://www.omnigroup.com/applications/omniweb/";
        _ver = detect_browser_ver("OmniWeb", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "omniweb";



    }
    else if (/Opera Mini/i.test(useragent))
    {
        _link = "http://www.opera.com/mini/";
        _ver = detect_browser_ver("Opera Mini", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "opera-2";



    }
    else if (/Opera Mobi/i.test(useragent))
    {
        _link = "http://www.opera.com/mobile/";
        _ver = detect_browser_ver("Opera Mobi", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "opera-2";



    }
    else if (/Opera Labs/i.test(useragent)
    || (/Opera/i.test(useragent)
    && /Edition Labs/i.test(useragent)))
    {
        _link = "http://labs.opera.com/";
        _ver = detect_browser_ver("Opera Labs", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "opera-next";



    }
    else if (/Opera Next/i.test(useragent)
    || (/Opera/i.test(useragent)
    && /Edition Next/i.test(useragent)))
    {
        _link = "http://www.opera.com/support/kb/view/991/";
        _ver = detect_browser_ver("Opera Next", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "opera-next";



    }
    else if (/Opera/i.test(useragent))
    {
        _link = "http://www.opera.com/";
        _ver = detect_browser_ver("Opera", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "opera-1";
        if (/_ver/i.test(useragent))
        code = "opera-2";



    }
    else if (/Orca/i.test(useragent))
    {
        _link = "http://www.orcabrowser.com/";
        _ver = detect_browser_ver("Orca", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "orca";



    }
    else if (/Oregano/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Oregano_(web_browser)";
        _ver = detect_browser_ver("Oregano", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "oregano";



    }
    else if (/Origyn\ Web\ Browser/i.test(useragent))
    {
        _link = "http://www.sand-labs.org/owb";
        title = "Oregano Web Browser";
        code = "owb";



    }
    else if (/osb-browser/i.test(useragent))
    {
        _link = "http://gtk-webcore.sourceforge.net/";
        _ver = detect_browser_ver("osb-browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/\ Pre\//i.test(useragent))
    {
        _link = "http://www.palm.com/us/products/phones/pre/index.html";
        _ver = detect_browser_ver("Pre", useragent);
        title = "Palm " + _ver.full;
        ver = _ver.version;
        code = "palmpre";



    }
    else if (/Palemoon/i.test(useragent))
    {
        _link = "http://www.palemoon.org/";
        _ver = detect_browser_ver("Moon", useragent);
        title = "Pale " + _ver.full;
        ver = _ver.version;
        code = "palemoon";



    }
    else if (/Patriott\:\:Browser/i.test(useragent))
    {
        _link = "http://madgroup.x10.mx/patriott1.php";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Patriott " + _ver.full;
        ver = _ver.version;
        code = "patriott";



    }
    else if (/Phaseout/i.test(useragent))
    {
        _link = "http://www.phaseout.net/";
        title = "Phaseout";
        code = "phaseout";



    }
    else if (/PhantomJS/i.test(useragent))
    {
        _link = "http://phantomjs.org/";
        _ver = detect_browser_ver("PhantomJS", useragent);
        title = _ver.full;
        ver = _ver.version;
        code = "phantomjs";



    }

    else if (/Phoenix/i.test(useragent))
    {
        _link = "http://www.mozilla.org/projects/phoenix/phoenix-release-notes.html";
        _ver = detect_browser_ver("Phoenix", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "phoenix";



    }
    else if (/Podkicker/i.test(useragent))
    {
        _link = "http://www.podkicker.com/";
        _ver = detect_browser_ver("Podkicker", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "podkicker";



    }
    else if (/Podkicker\ Pro/i.test(useragent))
    {
        _link = "http://www.podkicker.com/";
        _ver = detect_browser_ver("Podkicker Pro", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "podkicker";



    }
    else if (/Pogo/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/AT%26T_Pogo";
        _ver = detect_browser_ver("Pogo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "pogo";



    }
    else if (/Polaris/i.test(useragent))
    {
        _link = "http://www.infraware.co.kr/eng/01_product/product02.asp";
        _ver = detect_browser_ver("Polaris", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "polaris";



    }
    else if (/Prism/i.test(useragent))
    {
        _link = "http://prism.mozillalabs.com/";
        _ver = detect_browser_ver("Prism", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "prism";



    }
    else if (/QtWeb\ Internet\ Browser/i.test(useragent))
    {
        _link = "http://www.qtweb.net/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "QtWeb Internet " + _ver.full;
        ver = _ver.version;
        code = "qtwebinternetbrowser";



    }
    else if (/QupZilla/i.test(useragent))
    {
        _link = "http://www.qupzilla.com/";
        _ver = detect_browser_ver("QupZilla", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "qupzilla";



    }
    else if (/rekonq/i.test(useragent))
    {
        _link = "http://rekonq.sourceforge.net/";
        title = "rekonq";
        code = "rekonq";



    }
    else if (/retawq/i.test(useragent))
    {
        _link = "http://retawq.sourceforge.net/";
        _ver = detect_browser_ver("retawq", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "terminal";



    }
    else if (/RockMelt/i.test(useragent))
    {
        _link = "http://www.rockmelt.com/";
        _ver = detect_browser_ver("RockMelt", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "rockmelt";



    }
    else if (/Ryouko/i.test(useragent))
    {
        _link = "http://sourceforge.net/projects/ryouko/";
        _ver = detect_browser_ver("Ryouko", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ryouko";



    }
    else if (/SaaYaa/i.test(useragent))
    {
        _link = "http://www.saayaa.com/";
        title = "SaaYaa Explorer";
        code = "saayaa";



    }
    else if (/SeaMonkey/i.test(useragent))
    {
        _link = "http://www.seamonkey-project.org/";
        _ver = detect_browser_ver("SeaMonkey", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "seamonkey";



    }
    else if (/SEMC-Browser/i.test(useragent))
    {
        _link = "http://www.sonyericsson.com/";
        _ver = detect_browser_ver("SEMC-Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "semcbrowser";



    }
    else if (/SEMC-java/i.test(useragent))
    {
        _link = "http://www.sonyericsson.com/";
        _ver = detect_browser_ver("SEMC-java", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "semcbrowser";



    }
    else if (/Series60/i.test(useragent)
    && !/Symbian/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Web_Browser_for_S60";
        _ver = detect_browser_ver("Series60", useragent);
        title = "Nokia " + _ver.full;
        ver = _ver.version;
        code = "s60";



    }
    else if (/S60/i.test(useragent)
    && !/Symbian/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Web_Browser_for_S60";
        _ver = detect_browser_ver("S60", useragent);
        title = "Nokia " + _ver.full;
        ver = _ver.version;
        code = "s60";



    }
    else if (/SE\ /i.test(useragent)
    && /MetaSr/i.test(useragent))
    {
        _link = "http://ie.sogou.com/";
        title = "Sogou Explorer";
        code = "sogou";



    }
    else if (/Shiira/i.test(useragent))
    {
        _link = "http://www.shiira.jp/en.php";
        _ver = detect_browser_ver("Shiira", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "shiira";



    }
    else if (/Shiretoko/i.test(useragent))
    {
        _link = "http://www.mozilla.org/";
        _ver = detect_browser_ver("Shiretoko", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefoxdevpre";



    }
    else if (/Silk/i.test(useragent)
    && !/PlayStation/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Amazon_Silk";
        _ver = detect_browser_ver("Silk", useragent);
        title = "Amazon " + _ver.full;
        ver = _ver.version;

        code = "silk";



    }
    else if (/SiteKiosk/i.test(useragent))
    {
        _link = "http://www.sitekiosk.com/SiteKiosk/Default.aspx";
        _ver = detect_browser_ver("SiteKiosk", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "sitekiosk";



    }
    else if (/SkipStone/i.test(useragent))
    {
        _link = "http://www.muhri.net/skipstone/";
        _ver = detect_browser_ver("SkipStone", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "skipstone";



    }
    else if (/Skyfire/i.test(useragent))
    {
        _link = "http://www.skyfire.com/";
        _ver = detect_browser_ver("Skyfire", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "skyfire";



    }
    else if (/Sleipnir/i.test(useragent))
    {
        _link = "http://www.fenrir-inc.com/other/sleipnir/";
        _ver = detect_browser_ver("Sleipnir", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "sleipnir";



    }
    else if (/SlimBoat/i.test(useragent))
    {
        _link = "http://slimboat.com/";
        _ver = detect_browser_ver("SlimBoat", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "slimboat";



    }
    else if (/SlimBrowser/i.test(useragent))
    {
        _link = "http://www.flashpeak.com/sbrowser/";
        _ver = detect_browser_ver("SlimBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "slimbrowser";



    }
    else if (/SmartTV/i.test(useragent))
    {
        _link = "http://www.freethetvchallenge.com/details/faq";
        _ver = detect_browser_ver("SmartTV", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "maplebrowser";



    }
    else if (/Songbird/i.test(useragent))
    {
        _link = "http://www.getsongbird.com/";
        _ver = detect_browser_ver("Songbird", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "songbird";



    }
    else if (/Stainless/i.test(useragent))
    {
        _link = "http://www.stainlessapp.com/";
        _ver = detect_browser_ver("Stainless", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "stainless";



    }
    else if (/SubStream/i.test(useragent))
    {
        _link = "http://itunes.apple.com/us/app/substream/id389906706?mt=8";
        _ver = detect_browser_ver("SubStream", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "substream";



    }
    else if (/Sulfur/i.test(useragent))
    {
        _link = "http://www.flock.com/";
        _ver = detect_browser_ver("Sulfur", useragent);
        title = "Flock " + _ver.full;
        ver = _ver.version;
        code = "flock";



    }
    else if (/Sundance/i.test(useragent))
    {
        _link = "http://digola.com/sundance.html";
        _ver = detect_browser_ver("Sundance", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "sundance";



    }
    else if (/Sunrise/i.test(useragent))
    {
        _link = "http://www.sundialbrowser.com/";
        _ver = detect_browser_ver("Sundial", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "sundial";



    }
    else if (/Sunrise/i.test(useragent))
    {
        _link = "http://www.sunrisebrowser.com/";
        _ver = detect_browser_ver("Sunrise", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "sunrise";



    }
    else if (/Surf/i.test(useragent))
    {
        _link = "http://surf.suckless.org/";
        _ver = detect_browser_ver("Surf", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "surf";



    }
    else if (/Swiftfox/i.test(useragent))
    {
        _link = "http://www.getswiftfox.com/";
        _ver = detect_browser_ver("Swiftfox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "swiftfox";



    }
    else if (/Swiftweasel/i.test(useragent))
    {
        _link = "http://swiftweasel.tuxfamily.org/";
        _ver = detect_browser_ver("Swiftweasel", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "swiftweasel";



    }
    else if (/Sylera/i.test(useragent))
    {
        _link = "http://dombla.net/sylera/";
        _ver = detect_browser_ver("Sylera", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
	else if (/(ali|tao)browser/i.test(useragent))
    {
        _link = "http://browser.taobao.com/"; 
		_ver = detect_browser_ver("TaoBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "taobao";



    }

    else if (/tear/i.test(useragent))
    {
        _link = "http://wiki.maemo.org/Tear";
        title = "Tear";
        code = "tear";



    }
    else if (/TeaShark/i.test(useragent))
    {
        _link = "http://www.teashark.com/";
        _ver = detect_browser_ver("TeaShark", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "teashark";



    }
    else if (/Teleca/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/Obigo_Browser/";
        _ver = detect_browser_ver(" Teleca", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "obigo";



    }
    else if (/TencentTraveler/i.test(useragent))
    {
        _link = "http://www.tencent.com/en-us/index.shtml";
        title = "Tencent Traveler";
        code = "tencenttraveler";



    }
    else if (/TenFourFox/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/TenFourFox";
        _ver = detect_browser_ver("TenFourFox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "tenfourfox";



    }
    else if (/TheWorld/i.test(useragent))
    {
        _link = "http://www.ioage.com/";
        title = "TheWorld Browser";
        code = "theworld";



    }
    else if (/Thunderbird/i.test(useragent))
    {
        _link = "http://www.mozilla.com/thunderbird/";
        _ver = detect_browser_ver("Thunderbird", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "thunderbird";



    }
    else if (/Tizen/i.test(useragent))
    {
        _link = "https://www.tizen.org/";
        _ver = detect_browser_ver("Tizen", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "tizen";



    }
    else if (/Tjusig/i.test(useragent))
    {
        _link = "http://www.tjusig.cz/";
        _ver = detect_browser_ver("Tjusig", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "tjusig";



    }
    else if (/TencentTraveler/i.test(useragent))
    {
        _link = "http://tt.qq.com/";
        _ver = detect_browser_ver("TencentTraveler", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "tt-explorer";




    }
    else if (/uBrowser/i.test(useragent))
    {
        _link = "http://www.ubrowser.com/";
        _ver = detect_browser_ver("uBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ubrowser";



    }
    else if (/UC\ Browser/i.test(useragent))
    {
        _link = "http://www.uc.cn/English/index.shtml";
        _ver = detect_browser_ver("UC Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ucbrowser";



    }
    else if (/UCWEB/i.test(useragent))
    {
        _link = "http://www.ucweb.com/English/product.shtml";
        _ver = detect_browser_ver("UCWEB", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ucweb";



    }
    else if (/UltraBrowser/i.test(useragent))
    {
        _link = "http://www.ultrabrowser.com/";
        _ver = detect_browser_ver("UltraBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "ultrabrowser";



    }
    else if (/UP.Browser/i.test(useragent))
    {
        _link = "http://www.openwave.com/";
        _ver = detect_browser_ver("UP.Browser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "openwave";



    }
    else if (/UP._link/i.test(useragent))
    {
        _link = "http://www.openwave.com/";
        _ver = detect_browser_ver("UP._link", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "openwave";



    }
    else if (/Usejump/i.test(useragent))
    {
        _link = "http://www.usejump.com/";
        _ver = detect_browser_ver("Usejump", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "usejump";



    }
    else if (/uZardWeb/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/UZard_Web";
        _ver = detect_browser_ver("uZardWeb", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "uzardweb";



    }
    else if (/uZard/i.test(useragent))
    {
        _link = "http://en.wikipedia.org/wiki/UZard_Web";
        _ver = detect_browser_ver("uZard", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "uzardweb";



    }
    else if (/uzbl/i.test(useragent))
    {
        _link = "http://www.uzbl.org/";
        title = "uzbl";
        code = "uzbl";



    }
    else if (/Vimprobable/i.test(useragent))
    {
        _link = "http://www.vimprobable.org/";
        _ver = detect_browser_ver("Vimprobable", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/Vonkeror/i.test(useragent))
    {
        _link = "http://zzo38computer.cjb.net/vonkeror/";
        _ver = detect_browser_ver("Vonkeror", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/w3m/i.test(useragent))
    {
        _link = "http://w3m.sourceforge.net/";
        _ver = detect_browser_ver("W3M", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "w3m";



    }
    else if (/AppleWebkit/i.test(useragent)
    && /Android/i.test(useragent)
    && !/Chrome/i.test(useragent))
    {
        _link = "http://developer.android.com/reference/android/webkit/package-summary.html";
        _ver = detect_browser_ver("Android Webkit", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "android-webkit";



    }
    else if (/WebianShell/i.test(useragent))
    {
        _link = "http://webian.org/shell/";
        _ver = detect_browser_ver("Shell", useragent);
        title = "Webian " + _ver.full;
        ver = _ver.version;
        code = "webianshell";



    }
    else if (/Webrender/i.test(useragent))
    {
        _link = "http://webrender.99k.org/";
        title = "Webrender";
        code = "webrender";



    }
    else if (/WeltweitimnetzBrowser/i.test(useragent))
    {
        _link = "http://weltweitimnetz.de/software/Browser.en.page";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Weltweitimnetz " + _ver.full;
        ver = _ver.version;
        code = "weltweitimnetzbrowser";



    }
    else if (/wKiosk/i.test(useragent))
    {
        _link = "http://www.app4mac.com/store/index.php?target=products&product_id=9";
        title = "wKiosk";
        code = "wkiosk";



    }
    else if (/WorldWideWeb/i.test(useragent))
    {
        _link = "http://www.w3.org/People/Berners-Lee/WorldWideWeb.html";
        _ver = detect_browser_ver("WorldWideWeb", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "worldwideweb";



    }
    else if (/wp-android/i.test(useragent))
    {
        _link = "http://android.wordpress.org/";
        _ver = detect_browser_ver("wp-android", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/wp-blackberry/i.test(useragent))
    {
        _link = "http://blackberry.wordpress.org/";
        _ver = detect_browser_ver("wp-blackberry", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/wp-iphone/i.test(useragent))
    {
        _link = "http://ios.wordpress.org/";
        _ver = detect_browser_ver("wp-iphone", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/wp-nokia/i.test(useragent))
    {
        _link = "http://nokia.wordpress.org/";
        _ver = detect_browser_ver("wp-nokia", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/wp-webos/i.test(useragent))
    {
        _link = "http://webos.wordpress.org/";
        _ver = detect_browser_ver("wp-webos", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/wp-windowsphone/i.test(useragent))
    {
        _link = "http://windowsphone.wordpress.org/";
        _ver = detect_browser_ver("wp-windowsphone", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "wordpress";



    }
    else if (/Wyzo/i.test(useragent))
    {
        _link = "http://www.wyzo.com/";
        _ver = detect_browser_ver("Wyzo", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "Wyzo";



    }
    else if (/X-Smiles/i.test(useragent))
    {
        _link = "http://www.xsmiles.org/";
        _ver = detect_browser_ver("X-Smiles", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "x-smiles";



    }
    else if (/Xiino/i.test(useragent))
    {
        _link = "#";
        _ver = detect_browser_ver("Xiino", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "null";



    }
    else if (/YaBrowser/i.test(useragent))
    {
        _link = "http://browser.yandex.com/";
        _ver = detect_browser_ver("Browser", useragent);
        title = "Yandex." + _ver.full;
        ver = _ver.version;
        code = "yandex";



    }
    else if (/zBrowser/i.test(useragent))
    {
        _link = "http://sites.google.com/site/zeromusparadoxe01/zbrowser";
        _ver = detect_browser_ver("zBrowser", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "zbrowser";



    }
    else if (/ZipZap/i.test(useragent))
    {
        _link = "http://www.zipzaphome.com/";
        _ver = detect_browser_ver("ZipZap", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "zipzap";



    }

    // Pulled out of order to help ensure better detection for above browsers
    else if (/ABrowse/i.test(useragent))
    {
        _link = "http://abrowse.sourceforge.net/";
        _ver = detect_browser_ver("ABrowse", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "abrowse";



    }
    else if (/Chrome/i.test(useragent))
    {
        _link = "http://google.com/chrome/";
        _ver = detect_browser_ver("Chrome", useragent);
        title = (/chrome.+?mobile/i.test(useragent)?"Mobile ":"")+ "Google " + _ver.full;
        ver = _ver.version;
        code = "chrome";



    }
    else if (/Safari/i.test(useragent)
    && !/Nokia/i.test(useragent))
    {
        _link = "http://www.apple.com/safari/";
        title = "Safari";

        if (/_ver/i.test(useragent))
        {
            _ver = detect_browser_ver("Safari", useragent);
            ver = _ver.version;
            title = _ver.full;



        }

        if (/Mobile Safari/i.test(useragent))
        {
            title = "Mobile " + title;



        }

        code = "safari";



    }
    else if (/Nokia/i.test(useragent))
    {
        _link = "http://www.nokia.com/browser";
        title = "Nokia Web Browser";
        code = "maemo";



    }
    else if (/Firefox/i.test(useragent))
    {

        _link = "http://www.mozilla.org/";
        _ver = detect_browser_ver("Firefox", useragent);
        ver = _ver.version;
        title = _ver.full;
        code = "firefox";



    }
    else if (/MSIE/i.test(useragent))
    {
        _link = "http://www.microsoft.com/windows/products/winfamily/ie/default.mspx";
        _ver = detect_browser_ver("MSIE", useragent);
        title = "Internet Explorer" + _ver.full;
        ver = _ver.version;

        var regmatch=/MSIE[\ |\/]?([.0-9a-zA-Z]+)/i.exec(useragent);

        if (regmatch[1] >= 10)
        {
            code = "msie10";



        }
        else if (regmatch[1] >= 9)
        {
            code = "msie9";



        }
        else if (regmatch[1] >= 8)
        {
            code = "msie7";
		}
		else if (regmatch[1] >= 7)
		{
			/*var s = /Windows NT (6\.2|1)/i.exec(useragent);
			if (s.length == 0)
			{*/
				code = "msie7"
			/*}
			else
			{
				if (s[1] == "6.1")
				{
					code = "msie9";   //Windows 7 + Internet Explorer 9 in Compatibility Mode
					ver = "9";
					title = "Internet Explorer 9";

				}
				else
				{
					code = "msie10";  //Windows 8 + Internet Explorer 10 in Compatibility Mode
					ver = "10";
					title = "Internet Explorer 10";
					
				}
			}*/
			
        }
        else if (regmatch[1] >= 6)
        {
            code = "msie6";



        }
        else if (regmatch[1] >= 4)
        {
            // also ie5
            code = "msie4";



        }
        else if (regmatch[1] >= 3)
        {
            code = "msie3";



        }
        else if (regmatch[1] >= 2)
        {
            code = "msie2";



        }
        else if (regmatch[1] >= 1)
        {
            code = "msie1";



        }
        else
        {
            code = "msie";



        }



    }
    else if (/Mozilla/i.test(useragent))
    {
        _link = "http://www.mozilla.org/";
        title = "Mozilla Compatible";

        if (/rv:([.0-9a-zA-Z]+)/i.test(useragent)) {
            regmatch = /rv:([.0-9a-zA-Z]+)/i.exec(useragent);
            title = "Mozilla " + regmatch[1];



        }

        code = "mozilla";



    }
    else
    {
        _link = "#";
        title = "Unknown";
        code = "null";



    }

    var json = {
        "link": _link,
        "text": title,
        "filename": code,
        "ver": ver,
        "folder": "net",
        "fullfilename": "16/net/"+code+".png"



    }
    return json;



}
</script>